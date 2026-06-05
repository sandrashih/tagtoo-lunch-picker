const SHEET_NAME = 'restaurants'
const COL = { id:0, name:1, name_en:2, category:3, hours:4, url:5, lat:6, lng:7 }

function corsResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON)
}

function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action
    if (action === 'parse') {
      var url = e.parameter.url || ''
      var parsed = parseGoogleMapsUrl(url)
      return corsResponse(parsed || { error: 'cannot parse' })
    }
    return corsResponse(listRestaurants())
  } catch(err) { return corsResponse({ error: err.message }) }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents)
    if (body.action === 'add') return corsResponse(addRestaurant(body))
    return corsResponse({ status: 'error', message: 'unknown action' })
  } catch(err) { return corsResponse({ status: 'error', message: err.message }) }
}

function listRestaurants() {
  var sheet = getSheet()
  var rows = sheet.getDataRange().getValues()
  if (rows.length <= 1) return { restaurants: [] }
  var restaurants = rows.slice(1)
    .map(function(row, i) { return { id:String(i+2), name:row[COL.name]||'', name_en:row[COL.name_en]||'',
      category:row[COL.category]||'', hours:row[COL.hours]||'', url:row[COL.url]||'', lat:row[COL.lat]||'', lng:row[COL.lng]||'' } })
    .filter(function(r) { return r.name && r.lat && r.lng })
  return { restaurants: restaurants }
}

function addRestaurant(body) {
  var url = String(body.url||'').trim()
  var name = String(body.name||'').trim() || '未命名餐廳'
  var category = String(body.category||'').trim()
  var lat = parseFloat(body.lat), lng = parseFloat(body.lng)
  if (!url) return { status:'error', message:'missing url' }
  if (isNaN(lat)||isNaN(lng)) return { status:'error', message:'missing coordinates' }
  var sheet = getSheet(), rows = sheet.getDataRange().getValues()
  for (var i=1; i<rows.length; i++) {
    var eu=rows[i][COL.url], elat=rows[i][COL.lat], elng=rows[i][COL.lng]
    if (eu && eu===url) return {status:'duplicate'}
    if (elat && elng && Math.abs(Number(elat)-lat)<0.0002 && Math.abs(Number(elng)-lng)<0.0002) return {status:'duplicate'}
  }
  sheet.appendRow(['',name,'',category,'',url,lat,lng])
  return {status:'ok'}
}

function extractMapsUrl(html) {
  var idx = html.indexOf('google.com/maps/place/')
  if (idx === -1) return null
  var start = html.lastIndexOf('"', idx)
  var end = html.indexOf('"', idx)
  if (start === -1 || end === -1) return null
  return html.substring(start+1, end)
}

function parseGoogleMapsUrl(url) {
  var lat = null, lng = null, name = '', nameEn = ''

  var pm = url.match(/maps\/place\/([^\/@]+)\/@([\d.\-]+),([\d.\-]+)/)
  if (pm) { name=decodeURIComponent(pm[1].replace(/\+/g,' ')); lat=parseFloat(pm[2]); lng=parseFloat(pm[3]) }

  if (!lat) {
    var qm = url.match(/[?&](?:q|ll)=([\d.\-]+),([\d.\-]+)/)
    if (qm) { lat=parseFloat(qm[1]); lng=parseFloat(qm[2]) }
  }
  if (!lat) {
    var am = url.match(/@([\d.\-]+),([\d.\-]+)/)
    if (am) { lat=parseFloat(am[1]); lng=parseFloat(am[2]) }
  }

  if (!lat && (url.indexOf('goo.gl') !== -1 || url.indexOf('maps.app') !== -1)) {
    try {
      var opts1 = { followRedirects:false, muteHttpExceptions:true,
        headers:{'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'} }
      var r1 = UrlFetchApp.fetch(url, opts1)
      var loc = r1.getHeaders()['Location'] || r1.getHeaders()['location']
      if (loc) { var parsed1 = parseGoogleMapsUrl(loc); if (parsed1) return parsed1 }

      var opts2 = { followRedirects:true, muteHttpExceptions:true,
        headers:{'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'} }
      var r2 = UrlFetchApp.fetch(url, opts2)
      var finalUrl = r2.getUrl()
      if (finalUrl && finalUrl !== url) { var parsed2 = parseGoogleMapsUrl(finalUrl); if (parsed2) return parsed2 }

      var html = r2.getContentText()
      var mapsUrl = extractMapsUrl(html)
      if (mapsUrl) { var parsed3 = parseGoogleMapsUrl(mapsUrl); if (parsed3) return parsed3 }

      var cm = html.match(/!3d([\d.\-]+)!4d([\d.\-]+)/)
      if (cm) { lat=parseFloat(cm[1]); lng=parseFloat(cm[2]) }
      var tm = html.match(/<title>([^<]+?)(?:\s*[-–]\s*Google|<)/)
      if (tm && !name) name = tm[1].trim()
    } catch(e) {}
  }

  if (!lat) return null
  if (name) { nameEn = name.replace(/[^\x00-\x7F]/g,'').trim(); if(!nameEn) nameEn=name }
  return { name:name||'未命名餐廳', name_en:nameEn||'Unnamed', url:url, lat:lat, lng:lng }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) throw new Error('找不到工作表：'+SHEET_NAME)
  return sheet
}
