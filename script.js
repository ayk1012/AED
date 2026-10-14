'use strict';
// ===== 設定 =====
const DATA_FILE = 'AED.xlsx';
const CANDIDATES = 5;          // 直線距離の近い順に、実経路を調べる件数
const WALK_M_PER_MIN = 80;     // 徒歩速度(約4.8km/h)
const ROUTE_API = 'https://routing.openstreetmap.de/routed-foot/route/v1/foot/';

const map = L.map('map').setView([32.8032, 130.7079], 14);
L.tileLayer('https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png', {
  maxZoom: 18,
  attribution: '<a href="https://maps.gsi.go.jp/development/ichiran.html">地理院タイル</a>'
}).addTo(map);

let aeds = [], aedMarkers = [], startMarker = null, routeLayer = null;
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const setStatus = (msg, err = false) => { const s = $('status'); s.textContent = msg; s.className = err ? 'error' : ''; };
const fmtDist = m => m >= 1000 ? (m / 1000).toFixed(2) + ' km' : Math.round(m) + ' m';

// 2点間の直線距離(ハーバーサイン, m)
function haversine(a, b) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (b.lat - a.lat) * rad, dLng = (b.lng - a.lng) * rad;
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h));
}

// ===== Excel読み込み =====
async function loadAeds() {
  try {
    const res = await fetch(DATA_FILE);
    if (!res.ok) throw new Error(res.status);
    const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    aeds = rows.map(r => ({
      name: r['名称'], address: r['住所'], place: r['設置場所'], hours: r['利用可能時間'], note: r['備考'],
      lat: parseFloat(r['緯度']), lng: parseFloat(r['経度'])
    })).filter(a => a.name && isFinite(a.lat) && isFinite(a.lng));
    aeds.forEach(a => {
      const m = L.marker([a.lat, a.lng], { icon: L.divIcon({ className: '', html: '<div class="aed-icon">+</div>', iconSize: [24, 24] }) })
        .bindPopup(`<b>${esc(a.name)}</b><br>${esc(a.address)}<br>設置場所: ${esc(a.place)}<br>利用: ${esc(a.hours)}`).addTo(map);
      aedMarkers.push(m);
    });
    if (aeds.length) map.fitBounds(L.featureGroup(aedMarkers).getBounds().pad(0.2));
    setStatus(`AED ${aeds.length} 件を読み込みました。`);
  } catch (e) {
    setStatus('AED.xlsx を読み込めません。file:// ではなく、GitHub Pages等のHTTP経由で開いてください。', true);
  }
}

// ===== 出発地点 =====
function setStart(latlng) {
  if (startMarker) startMarker.setLatLng(latlng);
  else {
    startMarker = L.marker(latlng, { draggable: true, title: '出発地点(ドラッグで変更可)' }).addTo(map);
    startMarker.on('dragend', () => search(startMarker.getLatLng()));
  }
}
map.on('click', e => { setStart(e.latlng); search(e.latlng); });

$('btn-locate').addEventListener('click', () => {
  if (!navigator.geolocation) return setStatus('このブラウザは位置情報に対応していません。地図をタップして指定してください。', true);
  setStatus('現在地を取得中...');
  navigator.geolocation.getCurrentPosition(
    p => { const ll = L.latLng(p.coords.latitude, p.coords.longitude); setStart(ll); search(ll); },
    () => setStatus('現在地を取得できませんでした。地図をタップして出発地点を指定してください。', true),
    { enableHighAccuracy: true, timeout: 10000 }
  );
});

$('btn-clear').addEventListener('click', () => {
  if (routeLayer) map.removeLayer(routeLayer);
  if (startMarker) map.removeLayer(startMarker);
  routeLayer = startMarker = null;
  aedMarkers.forEach(m => m.getElement()?.firstChild.classList.remove('best'));
  $('result').hidden = true;
  setStatus('');
});

// ===== 最短経路探索 =====
async function fetchRoute(from, to) {
  const url = `${ROUTE_API}${from.lng},${from.lat};${to.lng},${to.lat}?overview=full&geometries=geojson`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(res.status);
  const j = await res.json();
  if (j.code !== 'Ok') throw new Error(j.code);
  return { dist: j.routes[0].distance, coords: j.routes[0].geometry.coordinates.map(c => [c[1], c[0]]) };
}

async function search(from) {
  if (!aeds.length) return setStatus('AEDデータがありません。', true);
  $('btn-locate').disabled = true;
  setStatus('最短経路を計算中...');
  // 1) 直線距離で候補を絞り込み → 2) 実際の徒歩経路距離で最短を決定
  const cands = aeds.map(a => ({ a, straight: haversine(from, a) })).sort((x, y) => x.straight - y.straight).slice(0, CANDIDATES);
  const results = await Promise.all(cands.map(async c => {
    try { return { ...c, ...(await fetchRoute(from, c.a)), real: true }; }
    catch { return { ...c, dist: c.straight, coords: [[from.lat, from.lng], [c.a.lat, c.a.lng]], real: false }; }
  }));
  results.sort((x, y) => x.dist - y.dist);
  const best = results[0];
  showResult(from, best, results.every(r => !r.real));
  $('btn-locate').disabled = false;
}

function showResult(from, best, fallback) {
  if (routeLayer) map.removeLayer(routeLayer);
  routeLayer = L.polyline(best.coords, { color: '#16794c', weight: 6, opacity: .85, dashArray: best.real ? null : '8 8' }).addTo(map);
  map.fitBounds(routeLayer.getBounds().pad(0.25));
  aedMarkers.forEach((m, i) => {
    const el = m.getElement()?.firstChild;
    if (el) el.classList.toggle('best', aeds[i] === best.a);
  });
  const a = best.a, min = Math.max(1, Math.round(best.dist / WALK_M_PER_MIN));
  const gm = `https://www.google.com/maps/dir/?api=1&origin=${from.lat},${from.lng}&destination=${a.lat},${a.lng}&travelmode=walking`;
  const r = $('result');
  r.hidden = false;
  r.innerHTML = `<h2>最寄りのAED: ${esc(a.name)}</h2>
    <div class="big">${fmtDist(best.dist)}・徒歩約${min}分</div>
    <dl><dt>住所</dt><dd>${esc(a.address)}</dd><dt>設置場所</dt><dd>${esc(a.place)}</dd>
    <dt>利用時間</dt><dd>${esc(a.hours)}</dd>${a.note ? `<dt>備考</dt><dd>${esc(a.note)}</dd>` : ''}</dl>
    <a class="gmap" href="${gm}" target="_blank" rel="noopener">Googleマップで道案内を開く ↗</a>`;
  setStatus(fallback ? '経路サービスに接続できないため、直線距離で表示しています。' : '緑の線が最短の徒歩経路です。');
}

loadAeds();
