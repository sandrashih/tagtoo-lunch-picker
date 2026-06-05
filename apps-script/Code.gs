// ===== 試算表欄位對應（依實際欄位順序）=====
// A: id  B: name  C: name_en  D: category  E: hours  F: url  G: lat  H: lng
const SHEET_NAME = 'restaurants'
const COL = { id:0, name:1, name_en:2, category:3, hours:4, url:5, lat:6, lng:7 }

// ===== CORS Helper =====
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
}

// ===== GET：列出所有餐廳 / 解析 URL =====
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action
    if (action === 'parse') {
      const url = e.parameter.url || ''
      const parsed = parseGoogleMapsUrl(url)
      return corsResponse(parsed || { error: 'cannot parse' })
    }
    return corsResponse(listRestaurants())
  } catch(err) {
    return corsResponse({ error: err.message })
  }
}

// ===== POST：新增餐廳 =====
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents)
    if (body.action === 'add') return corsResponse(addRestaurant(body))
    return corsResponse({ status: 'error', message: 'unknown action' })
  } catch(err) {
    return corsResponse({ status: 'error', message: err.message })
  }
}

// ===== 列出所有餐廳 =====
function listRestaurants() {
  const sheet = getSheet()
  const rows = sheet.getDataRange().getValues()
  if (rows.length <= 1) return { restaurants: [] }

  const restaurants = rows.slice(1)
    .map((row, i) => ({
      id:       String(i + 2),
      name:     row[COL.name]     || '',
      name_en:  row[COL.name_en]  || '',
      category: row[COL.category] || '',
      hours:    row[COL.hours]    || '',
      url:      row[COL.url]      || '',
      lat:      row[COL.lat]      || '',
      lng:      row[COL.lng]      || '',
    }))
    .filter(r => r.name && r.lat && r.lng)

  return { restaurants }
}

// ===== 新增餐廳（接收前端傳來的 name, url, lat, lng）=====
function addRestaurant(body) {
  const url      = String(body.url      || '').trim()
  const name     = String(body.name     || '').trim() || '未命名餐廳'
  const category = String(body.category || '').trim()
  const lat      = parseFloat(body.lat)
  const lng      = parseFloat(body.lng)

  if (!url) return { status: 'error', message: 'missing url' }
  if (isNaN(lat) || isNaN(lng)) return { status: 'error', message: 'missing coordinates' }

  const sheet = getSheet()
  const rows = sheet.getDataRange().getValues()

  // 重複檢查：座標（0.0002° ≈ 22m）或相同 URL
  for (let i = 1; i < rows.length; i++) {
    const eu = rows[i][COL.url], elat = rows[i][COL.lat], elng = rows[i][COL.lng]
    if (eu && eu === url) return { status: 'duplicate' }
    if (elat && elng &&
        Math.abs(Number(elat) - lat) < 0.0002 &&
        Math.abs(Number(elng) - lng) < 0.0002) {
      return { status: 'duplicate' }
    }
  }

  // id留空, name, name_en, category, hours, url, lat, lng
  sheet.appendRow(['', name, '', category, '', url, lat, lng])
  return { status: 'ok' }
}

// ===== 解析 Google Maps URL =====
function parseGoogleMapsUrl(url) {
  let lat = null, lng = null, name = '', nameEn = ''

  // 格式 1: maps/place/NAME/@lat,lng
  const placeMatch = url.match(/maps\/place\/([^/@]+)\/@([-\d.]+),([-\d.]+)/)
  if (placeMatch) {
    name = decodeURIComponent(placeMatch[1].replace(/\+/g, ' '))
    lat  = parseFloat(placeMatch[2])
    lng  = parseFloat(placeMatch[3])
  }

  // 格式 2: ?q=lat,lng 或 ll=lat,lng
  if (!lat) {
    const qMatch = url.match(/[?&](?:q|ll)=([-\d.]+),([-\d.]+)/)
    if (qMatch) { lat = parseFloat(qMatch[1]); lng = parseFloat(qMatch[2]) }
  }

  // 格式 3: @lat,lng
  if (!lat) {
    const atMatch = url.match(/@([-\d.]+),([-\d.]+)/)
    if (atMatch) { lat = parseFloat(atMatch[1]); lng = parseFloat(atMatch[2]) }
  }

  // 短網址展開（maps.app.goo.gl 返回 HTML，需從內容抓真實 URL）
  if (!lat && (url.includes('goo.gl') || url.includes('maps.app'))) {
    try {
      const resp = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true })
      const finalUrl = resp.getUrl()
      const html = resp.getContentText()

      // 先試最終 URL
      if (finalUrl && finalUrl !== url) {
        const r = parseGoogleMapsUrl(finalUrl)
        if (r) return r
      }

      // 從 HTML 裡找 Google Maps place URL
      const patterns = [
        /href="(https:\/\/www\.google\.com\/maps\/place\/[^"]+)"/,
        /content="(https:\/\/www\.google\.com\/maps\/place\/[^"]+)"/,
        /"(https:\/\/www\.google\.com\/maps\/place\/[^"]+)"/
      ]
      for (const pat of patterns) {
        const m = html.match(pat)
        if (m) {
          const r = parseGoogleMapsUrl(m[1])
          if (r) return r
        }
      }
    } catch(e) {}
  }

  if (!lat) return null

  // 嘗試提取英文名
  if (name) {
    nameEn = name.replace(/[^\x00-\x7F]/g, '').trim()
    if (!nameEn) nameEn = name
  }

  return { name: name || '未命名餐廳', name_en: nameEn || 'Unnamed', url, lat, lng }
}

// ===== 工具函式 =====
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME)
  return sheet
}
