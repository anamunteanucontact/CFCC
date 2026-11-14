/* ===================== HARTA COPIILOR (/hartacopiilor) =====================
 * Datele stau în „Comunitatea CFCC”, fila „Hartă” (listele în „Liste hartă”).
 * Coloane: A Status · B Nume · C Categorie · D Tip · E Zonă · F Adresă · G Lat · H Lng ·
 *          I Partener (ID din Parteneri) · J WhatsApp · K Verificat la · L Sursă.
 * Site-ul arată doar rândurile cu Status = Publicat și cu Lat/Lng completate.
 * Fișier separat în proiectul Apps Script (lângă formulare.gs). Nu atinge rândurile existente,
 * cu excepția celulelor goale Lat, Lng și Zonă.
 * O singură dată, din editor: rulează porneste_harta (completează coordonatele în fiecare noapte, la 5:00)
 * și harta_import_osm (aduce creșele, grădinițele și locurile de joacă din OpenStreetMap). */
const HARTA = { fila: 'Hartă', liste: 'Liste hartă', bbox: [45.68, 21.08, 45.83, 21.36] };
const HARTA_COMUNE = { 'dumbravita': 'Dumbrăvița', 'ghiroda': 'Ghiroda', 'giarmata vii': 'Ghiroda',
  'giroc': 'Giroc - Chișoda', 'chisoda': 'Giroc - Chișoda', 'mosnita noua': 'Moșnița Nouă',
  'mosnita veche': 'Moșnița Nouă', 'sacalaz': 'Săcălaz' };

function harta_norm_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

/** Zonele din „Liste hartă” (C) cu centrul lor aproximativ (D: „lat, lng”). */
function harta_zone_(com) {
  const f = com.getSheetByName(HARTA.liste);
  const v = f.getRange(2, 3, Math.max(f.getLastRow() - 1, 1), 2).getValues();
  return v.map(function (r) {
    const m = String(r[1]).match(/(-?\d+(?:\.\d+)?)\s*[,;]\s*(-?\d+(?:\.\d+)?)/);
    return m && r[0] ? { n: String(r[0]).trim(), lat: +m[1], lng: +m[2] } : null;
  }).filter(Boolean);
}

/** Zona: comuna din adresă, dacă e una din cele știute; altfel cartierul cu centrul cel mai apropiat. */
function harta_zona_(zone, lat, lng, adresa) {
  const a = harta_norm_(adresa);
  for (const k in HARTA_COMUNE) if (a.indexOf(k) >= 0) return HARTA_COMUNE[k];
  if (!lat || !lng) return '';
  let best = null, d = 1e9;
  zone.forEach(function (z) {
    const dx = (z.lng - lng) * Math.cos(lat * Math.PI / 180), dy = z.lat - lat, x = dx * dx + dy * dy;
    if (x < d) { d = x; best = z; }
  });
  return best ? best.n : '';
}

/** Adaugă din OpenStreetMap creșe, grădinițe și locuri de joacă care nu sunt încă în fila Hartă.
 *  Rândurile noi intră cu Status = De verificat (nu apar pe site până nu le pui Publicat). */
function harta_import_osm() {
  const com = SpreadsheetApp.openById(COMUNITATE_ID), fila = com.getSheetByName(HARTA.fila);
  const b = HARTA.bbox.join(',');
  const q = '[out:json][timeout:60];(' +
    'nwr["amenity"~"^(kindergarten|childcare)$"](' + b + ');' +
    'nwr["leisure"~"^(playground|indoor_play)$"](' + b + ');' +
    ');out center tags;';
  const r = UrlFetchApp.fetch('https://overpass-api.de/api/interpreter', {
    method: 'post', payload: { data: q }, muteHttpExceptions: true,
    headers: { 'User-Agent': 'cefacemcucopiii.ro (hello@cefacemcucopiii.ro)' } });
  if (r.getResponseCode() !== 200) throw new Error('OpenStreetMap nu a răspuns (' + r.getResponseCode() + '). Mai încearcă peste câteva minute.');
  const el = JSON.parse(r.getContentText()).elements || [];
  const ultim = fila.getLastRow();
  const exist = {};
  if (ultim > 1) fila.getRange(2, 12, ultim - 1, 1).getValues().forEach(function (x) { if (x[0]) exist[String(x[0]).trim()] = 1; });
  const zone = harta_zone_(com), rows = [];
  el.forEach(function (e) {
    const id = 'osm:' + e.type + '/' + e.id;
    if (exist[id]) return;
    const t = e.tags || {}, lat = e.lat || (e.center && e.center.lat), lng = e.lon || (e.center && e.center.lon);
    if (!lat || !lng) return;
    let cat, tip = '', nume = (t.name || '').trim();
    if (t.leisure) {
      cat = 'Loc de joacă';
      tip = (t.leisure === 'indoor_play' || t.indoor === 'yes') ? 'Indoor' : 'Outdoor';
      if (!nume) nume = tip === 'Indoor' ? 'Loc de joacă indoor' : 'Loc de joacă';
    } else {
      if (!nume) return; // creșe și grădinițe fără nume nu sunt utile pe hartă
      const n = harta_norm_(nume), cresa = t.amenity === 'childcare' || /cres[ae]/.test(n), grad = /gradinit/.test(n);
      cat = cresa && grad ? 'Creșă, Grădiniță' : cresa ? 'Creșă' : 'Grădiniță';
      const op = harta_norm_(t['operator:type']);
      tip = /public|government|municipal|state/.test(op) || /\bpp\b|program (prelungit|normal)|nr\.?\s*\d/.test(n) ? 'De stat'
        : /private|privat/.test(op) ? 'Privat' : '';
    }
    const strada = [t['addr:street'], t['addr:housenumber']].filter(Boolean).join(' ');
    const loc = t['addr:city'] || t['addr:place'] || '';
    const adresa = [strada, loc || (strada ? 'Timișoara' : '')].filter(Boolean).join(', ');
    rows.push(['De verificat', nume, cat, tip, harta_zona_(zone, lat, lng, loc || adresa), adresa,
      Math.round(lat * 1e6) / 1e6, Math.round(lng * 1e6) / 1e6, '', '', '', id]);
  });
  if (rows.length) fila.getRange(ultim + 1, 1, rows.length, 12).setValues(rows);
  harta_mesaj_(rows.length
    ? 'Am adăugat ' + rows.length + ' locații noi în fila Hartă, cu Status „De verificat”. Sursa: OpenStreetMap.'
    : 'Nu am găsit locații noi în OpenStreetMap.', 'Hartă', 10);
}

/** Completează Lat/Lng din Adresă (OpenStreetMap) și Zona, doar unde sunt goale.
 *  Câte o adresă pe secundă (regula OpenStreetMap); dacă sunt multe, mai pornești o dată. */
function harta_completeaza() {
  const com = SpreadsheetApp.openById(COMUNITATE_ID), fila = com.getSheetByName(HARTA.fila);
  const ultim = fila.getLastRow();
  if (ultim < 2) return;
  const v = fila.getRange(2, 1, ultim - 1, 8).getValues(), zone = harta_zone_(com), b = HARTA.bbox;
  const start = Date.now();
  let coord = 0, zon = 0, gresit = [], ramase = 0;
  for (let i = 0; i < v.length; i++) {
    const rand = i + 2, adresa = String(v[i][5]).trim();
    let lat = v[i][6], lng = v[i][7];
    if ((lat === '' || lng === '') && adresa) {
      if (Date.now() - start > 300000) { ramase++; continue; }
      const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=ro&bounded=1' +
        '&viewbox=' + [b[1], b[2], b[3], b[0]].join(',') +
        '&q=' + encodeURIComponent(/timi[sș]oara|dumbr|ghiroda|giroc|chi[sș]oda|mo[sș]ni|s[aă]c[aă]laz|giarmata/i.test(adresa) ? adresa : adresa + ', Timișoara');
      const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true,
        headers: { 'User-Agent': 'cefacemcucopiii.ro (hello@cefacemcucopiii.ro)' } });
      Utilities.sleep(1100);
      const x = r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : [];
      if (x.length) {
        lat = Math.round(+x[0].lat * 1e6) / 1e6; lng = Math.round(+x[0].lon * 1e6) / 1e6;
        fila.getRange(rand, 7, 1, 2).setValues([[lat, lng]]); coord++;
      } else { gresit.push(rand); continue; }
    }
    if (!String(v[i][4]).trim() && lat !== '' && lng !== '') {
      const z = harta_zona_(zone, +lat, +lng, adresa);
      if (z) { fila.getRange(rand, 5).setValue(z); zon++; }
    }
  }
  harta_mesaj_('Coordonate completate: ' + coord + ' · zone completate: ' + zon +
    (gresit.length ? ' · adrese negăsite (rândurile ' + gresit.slice(0, 15).join(', ') + (gresit.length > 15 ? '...' : '') + ') - scrie coordonatele de mână' : '') +
    (ramase ? ' · au mai rămas ' + ramase + ', mai pornește o dată' : ''), 'Hartă', 15);
}

/** Mesaj în sheet dacă e deschis; altfel doar în jurnalul execuției (rulare automată). */
function harta_mesaj_(text, titlu) {
  console.log(text);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(text, titlu, 15); } catch (e) {}
}

/** O singură dată, din editor: completarea coordonatelor rulează singură în fiecare noapte, la 5:00. */
function porneste_harta() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'harta_completeaza') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('harta_completeaza').timeBased().everyDays(1).atHour(5).create();
  harta_completeaza();
}
