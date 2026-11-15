/* ===================== HARTA COPIILOR (/hartacopiilor) =====================
 * Datele stau în „Comunitatea CFCC”, fila „Hartă” (listele în „Liste hartă”).
 * Coloane: A Status · B Nume · C Categorie · D Tip · E Zonă · F Adresă · G Lat · H Lng ·
 *          I Partener (ID din Parteneri) · J WhatsApp · K Verificat la · L Sursă.
 * Site-ul arată doar rândurile cu Status = Publicat și cu Lat/Lng completate.
 * Fișier separat în proiectul Apps Script (lângă formulare.gs). Nu atinge rândurile existente,
 * cu excepția celulelor goale Lat, Lng și Zonă.
 * O singură dată, din editor: rulează porneste_harta (completează coordonatele în fiecare noapte, la 5:00)
 * și harta_import_osm (aduce creșele, grădinițele și locurile de joacă din OpenStreetMap). */
// bbox = Timișoara + zona periurbană (Sânandrei, Giarmata, Remetea Mare, Urseni, Sânmihaiu Român)
const HARTA = { fila: 'Hartă', liste: 'Liste hartă', bbox: [45.66, 21.04, 45.88, 21.42] };
// Localitățile periurbane și zona lor pe hartă. Ordinea contează: „giarmata-vii” înainte de „giarmata”.
const HARTA_COMUNE = { 'dumbravita': 'Dumbrăvița', 'ghiroda': 'Ghiroda',
  'giarmata-vii': 'Giarmata-Vii', 'giarmata vii': 'Giarmata-Vii', 'giarmata': 'Giarmata', 'remetea mare': 'Remetea Mare',
  'giroc': 'Giroc - Chișoda', 'chisoda': 'Giroc - Chișoda',
  'mosnita noua': 'Moșnița Nouă', 'mosnita veche': 'Moșnița Veche', 'urseni': 'Urseni',
  'sacalaz': 'Săcălaz', 'sanmihaiu roman': 'Sânmihaiu Român', 'utvin': 'Sânmihaiu Român', 'sanandrei': 'Sânandrei' };

function harta_norm_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
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

/** Zona: comuna din adresă (doar localitatea, după ultima virgulă), altfel cartierul cu centrul cel mai apropiat.
 *  Așa „Calea Girocului, Timișoara” rămâne în Timișoara, nu în Giroc. */
function harta_zona_(zone, lat, lng, adresa) {
  const a = harta_norm_(String(adresa || '').split(',').pop());
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
  // Serverul principal (Germania), apoi două copii de rezervă (Austria, Germania), dacă primul e ocupat.
  const servere = ['https://overpass-api.de', 'https://overpass.private.coffee', 'https://overpass.kumi.systems'];
  let el = null, erori = [];
  for (let i = 0; i < servere.length && !el; i++) {
    try {
      const r = UrlFetchApp.fetch(servere[i] + '/api/interpreter', {
        method: 'post', payload: { data: q }, muteHttpExceptions: true,
        headers: { 'User-Agent': 'cefacemcucopiii.ro (hello@cefacemcucopiii.ro)' } });
      const t = r.getContentText();
      if (r.getResponseCode() === 200 && t.charAt(0) === '{') el = JSON.parse(t).elements || [];
      else erori.push(servere[i] + ': ' + r.getResponseCode());
    } catch (e) { erori.push(servere[i] + ': ' + e.message); }
  }
  if (!el) throw new Error('OpenStreetMap nu a răspuns (' + erori.join(' · ') + '). Mai încearcă peste câteva minute.');
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
      // În aer liber = „Parc”; „Loc de joacă” = locurile de joacă indoor, private (decizia Anei, 15 nov)
      tip = (t.leisure === 'indoor_play' || t.indoor === 'yes') ? 'Indoor' : 'Outdoor';
      cat = tip === 'Indoor' ? 'Loc de joacă' : 'Parc';
      if (!nume) nume = cat;
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
    const adresa = [strada, loc].filter(Boolean).join(', '); // fără adresă: o completează harta_completeaza din coordonate
    rows.push(['De verificat', nume, cat, tip, harta_zona_(zone, lat, lng, loc), adresa,
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
  const v = fila.getRange(2, 1, ultim - 1, 12).getValues(), zone = harta_zone_(com), b = HARTA.bbox;
  const start = Date.now();
  let coord = 0, zon = 0, adr = 0, gresit = [], ramase = 0;
  const dubluri = harta_dubluri_(v);
  dubluri.forEach(function (i) { fila.getRange(i + 2, 1).setValue('Ascuns'); v[i][0] = 'Ascuns'; });
  SpreadsheetApp.flush();
  // Google oprește orice rulare după 6 minute și pierde ce nu e salvat: salvăm la fiecare 10 rânduri
  // și ne oprim singuri după 4 minute. Ce rămâne se face la rularea următoare (sau noaptea, la 5:00).
  for (let i = 0; i < v.length; i++) {
    if (i % 10 === 0) SpreadsheetApp.flush();
    const rand = i + 2, adresa = String(v[i][5]).trim();
    let lat = v[i][6], lng = v[i][7];
    if ((lat === '' || lng === '') && adresa) {
      if (Date.now() - start > 240000) { ramase++; continue; }
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
    // Fără adresă, dar cu coordonate: adresa din coordonate (doar rândurile încă nepublicate).
    let adresaNoua = adresa;
    if (!adresa && lat !== '' && lng !== '' && v[i][0] !== 'Ascuns') {
      if (Date.now() - start > 240000) { ramase++; continue; }
      const a = harta_adresa_(+String(lat).replace(',', '.'), +String(lng).replace(',', '.'));
      if (a) {
        adresaNoua = a.adresa;
        fila.getRange(rand, 6).setValue(a.adresa); adr++;
        const nume = String(v[i][1]).trim();
        if (/^(parc|loc de joac[aă]|loc de joac[aă] indoor|cre[sș][aă]|gr[aă]dini[tț][aă])$/i.test(nume) && a.strada)
          fila.getRange(rand, 2).setValue(nume + ' - ' + a.strada);
        if (a.localitate) { const z0 = harta_zona_(zone, null, null, a.localitate); if (z0) { fila.getRange(rand, 5).setValue(z0); v[i][4] = z0; } }
      }
    }
    if (!String(v[i][4]).trim() && lat !== '' && lng !== '') {
      const z = harta_zona_(zone, +String(lat).replace(',', '.'), +String(lng).replace(',', '.'), adresaNoua);
      if (z) { fila.getRange(rand, 5).setValue(z); zon++; }
    }
  }
  SpreadsheetApp.flush();
  harta_mesaj_('Coordonate completate: ' + coord + ' · adrese din coordonate: ' + adr + ' · zone completate: ' + zon +
    (dubluri.length ? ' · dubluri ascunse: ' + dubluri.length : '') +
    (gresit.length ? ' · adrese negăsite (rândurile ' + gresit.slice(0, 15).join(', ') + (gresit.length > 15 ? '...' : '') + ') - scrie coordonatele de mână' : '') +
    (ramase ? ' · au mai rămas ' + ramase + ' - rulează harta_completeaza încă o dată (sau se fac singure la noapte)' : ''), 'Hartă', 15);
}

/** Adresa unui punct (OpenStreetMap): { adresa: 'Strada X 5, Timișoara', strada: 'Strada X', localitate: 'Timișoara' }. */
function harta_adresa_(lat, lng) {
  const r = UrlFetchApp.fetch('https://nominatim.openstreetmap.org/reverse?format=json&zoom=18&accept-language=ro&lat=' + lat + '&lon=' + lng,
    { muteHttpExceptions: true, headers: { 'User-Agent': 'cefacemcucopiii.ro (hello@cefacemcucopiii.ro)' } });
  Utilities.sleep(1100);
  if (r.getResponseCode() !== 200) return null;
  const a = (JSON.parse(r.getContentText()) || {}).address || {};
  const strada = a.road || a.pedestrian || a.footway || a.park || a.square || '';
  const localitate = a.city || a.town || a.village || a.municipality || '';
  if (!strada && !localitate) return null;
  return { adresa: [[strada, a.house_number].filter(Boolean).join(' '), localitate].filter(Boolean).join(', '), strada: strada, localitate: localitate };
}

/** Locurile de joacă fără nume, din OpenStreetMap, aflate la mai puțin de 40 m de alt loc de joacă din listă,
 *  sunt de obicei bucăți din același loc. Le ascundem (Status = Ascuns), doar dacă sunt încă „De verificat”. */
function harta_dubluri_(v) {
  const out = [], pastrate = [];
  const num = function (x) { return +String(x).replace(',', '.'); };
  for (let i = 0; i < v.length; i++) {
    const cat = String(v[i][2]), lat = num(v[i][6]), lng = num(v[i][7]);
    if ((cat.indexOf('Loc de joac') < 0 && cat.indexOf('Parc') < 0) || !lat || !lng || v[i][0] === 'Ascuns') continue;
    const faraNume = /^(parc|loc de joac[aă]( indoor)?)$/i.test(String(v[i][1]).trim());
    const aproape = pastrate.some(function (p) {
      const dx = (p[1] - lng) * 111320 * Math.cos(lat * Math.PI / 180), dy = (p[0] - lat) * 110540;
      return dx * dx + dy * dy < 1600;
    });
    if (aproape && faraNume && v[i][0] === 'De verificat' && /^osm:/.test(String(v[i][11]))) out.push(i);
    else pastrate.push([lat, lng]);
  }
  return out;
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

/* ===================== FORMULARUL „ADAUG UN LOC PE HARTĂ” (Colaborări, tip „loc”) =====================
 * Formularul de pe site trimite { tip: 'loc', campuri } la același doPost din Code.gs. Aici doar adăugăm tipul nou
 * în listele din Code.gs (FILE, NOTIF, TEXTE_REZERVA, ETICHETE_CONFIRMARE), ca să nu modificăm fișierul acela.
 * Răspunsurile intră în „CFCC - Onboarding parteneri” → fila „Propuneri locuri hartă” (se creează singură),
 * NU în fila Hartă: acolo sheet-ul e public, iar datele de contact ale celui care trimite nu au voie să ajungă public.
 * Atenție: formularul merge pe site abia după Deploy → Manage deployments → Edit → New version. */
FILE.loc = {
  nume: 'Propuneri locuri hartă', fisier: 'onboarding', status: 'Nou', rosu: true,
  coloane: [['Primit la'], ['Status'], ['Nume loc', 'Nume'], ['Ce fel de loc'], ['Adresă'], ['Website / Instagram'], ['Despre loc'],
    ['Cine trimite'], ['Persoană de contact'], ['E-mail'], ['Telefon'], ['Vrea să fie partener'], ['Note']]
};
NOTIF.loc = ['Loc propus pentru hartă', function (c) { return c['Nume'] || ''; },
  ['Nume', 'Ce fel de loc', 'Adresă', 'Cine trimite', 'Vrea să fie partener', 'Persoană de contact', 'E-mail', 'Telefon'],
  'CFCC - Onboarding parteneri → Propuneri locuri hartă', ONBOARDING_ID];
TEXTE_REZERVA.loc = ['Am primit locul propus: {pagina}',
  'Bună!\n\nMulțumim că ne-ai scris despre {pagina}. Îl verificăm și, dacă se potrivește, îl adăugăm pe Harta copiilor de pe {site}.\n\n' +
  'Dacă ai întrebări sau vrei să ne spui ceva în plus, răspunde la acest e-mail.\n\nCu drag,\nEchipa cefacemcucopiii.ro'];
ETICHETE_CONFIRMARE.loc = 'am primit locul';
CAND_PLEACA.loc = 'Cineva trimite un loc pentru Harta copiilor (Colaborări → Adaug un loc pe hartă)';
