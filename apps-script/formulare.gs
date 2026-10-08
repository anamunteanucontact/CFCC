/**
 * Ce facem cu copiii? - scriptul din sheet-ul principal (centralizatorul de evenimente).
 *
 * 1. Primește formularele de pe site:
 *    - propunerile de evenimente -> fila „Propuneri evenimente” din acest sheet
 *    - abonările și colaborările -> sheet-ul separat „Contacte”
 *    și trimite un e-mail de confirmare celui care a completat.
 * 2. Când schimbi statusul unei propuneri în „Pregătit pentru calendar”,
 *    o mută singur în fila „Evenimente”, la ziua și ora ei, cu status „De verificat”.
 *    După ce o treci pe „Confirmat”, apare pe site.
 */

const NUME_EXPEDITOR = 'Ce facem cu copiii?';
const SITE = 'https://cefacemcucopiii.ro';
const INSTAGRAM = 'https://www.instagram.com/cefacemcucopiii/';
const CONTACTE_ID = '1ydRGLLcpV0bGP07bckIoAZ4k2Z6llGmozHPJTVbSF38';

// Fiecare coloană: [titlul din sheet, numele câmpului trimis de site]
const FILE = {
  eveniment: {
    nume: 'Propuneri evenimente', fisier: 'principal',
    coloane: [['Primit la'], ['Nume eveniment'], ['Organizator'], ['Data'], ['Ora'], ['Locație'], ['Adresă'],
      ['Vârstă minimă'], ['Vârstă maximă'], ['Acces'], ['Link bilete / înscriere'], ['Descriere'], ['E-mail'], ['Telefon'], ['Status']]
  },
  colaborare: {
    nume: 'Colaborări', fisier: 'contacte',
    coloane: [['Primit la'], ['Afacere / brand'], ['Persoană de contact'], ['E-mail'], ['Telefon'], ['Instagram'],
      ['Facebook'], ['Website'], ['Despre ei'], ['Ce au nevoie'], ['Status'], ['Note']]
  },
  abonare: {
    nume: 'Abonați', fisier: 'contacte',
    coloane: [['Primit la'], ['E-mail'], ['Vrea oferte speciale', 'Vrea oferte de la parteneri'], ['Note']]
  }
};

/* ---------- 1. formularele de pe site ---------- */

function doPost(e) {
  try {
    const date = JSON.parse(e.postData.contents);
    const tip = FILE[date.tip];
    if (!tip) return raspuns({ ok: false, eroare: 'tip necunoscut' });
    const campuri = date.campuri || {};

    const fila = ia_fila(tip);
    const rand = tip.coloane.map(function (c) {
      const titlu = c[0], cheie = c[1] || c[0];
      if (titlu === 'Primit la') return new Date();
      if (titlu === 'Status') return 'Nou';
      const v = campuri[cheie];
      if (v === undefined || v === null || v === '') return '';
      // apostroful păstrează textul exact cum a fost scris (telefoane cu 0 în față, date, ore)
      return "'" + String(v).slice(0, 2000);
    });
    fila.appendRow(rand);

    try { trimite_confirmare(date.tip, campuri); } catch (errMail) { /* rândul e salvat oricum */ }
    return raspuns({ ok: true });
  } catch (err) {
    return raspuns({ ok: false, eroare: String(err) });
  }
}

function ia_fila(tip) {
  const ss = tip.fisier === 'contacte' ? SpreadsheetApp.openById(CONTACTE_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let fila = ss.getSheetByName(tip.nume);
  if (!fila) {
    fila = ss.insertSheet(tip.nume);
    fila.appendRow(tip.coloane.map(function (c) { return c[0]; }));
    fila.setFrozenRows(1);
    fila.getRange(1, 1, 1, tip.coloane.length).setFontWeight('bold');
  }
  return fila;
}

function raspuns(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- e-mailurile de confirmare ---------- */

function trimite_confirmare(tip, c) {
  const catre = String(c['E-mail'] || '').trim();
  if (!/^\S+@\S+\.\S+$/.test(catre)) return;
  let subiect, continut;

  if (tip === 'abonare') {
    subiect = 'Bine ai venit la Ce facem cu copiii?';
    continut =
      '<p>Ceau!</p>' +
      '<p>Mulțumim că te-ai abonat. De acum, în fiecare săptămână îți trimitem programul: spectacole, ateliere, concerte și ieșiri pentru copii din Timișoara.</p>' +
      (c['Vrea oferte de la parteneri'] === 'Da' ? '<p>Ai ales să primești și reduceri și oferte speciale pentru familii, de la organizatorii și brandurile cu care colaborăm. Le trimitem doar când merită.</p>' : '') +
      '<p>Până la primul e-mail, tot calendarul e pe <a href="' + SITE + '">cefacemcucopiii.ro</a>, iar noutățile zilnice pe <a href="' + INSTAGRAM + '">Instagram</a>.</p>' +
      '<p>Cu drag,<br>Ana, de la Ce facem cu copiii?</p>' +
      '<p style="color:#888;font-size:12px">Te poți dezabona oricând, răspunzând la acest e-mail cu „dezabonare”.</p>';
  } else if (tip === 'eveniment') {
    subiect = 'Am primit evenimentul tău: ' + (c['Nume eveniment'] || '');
    continut =
      '<p>Bună!</p>' +
      '<p>Mulțumim că ne-ai trimis <strong>' + esc(c['Nume eveniment']) + '</strong>. Îl verificăm și îl adăugăm în calendarul de pe <a href="' + SITE + '">cefacemcucopiii.ro</a>.</p>' +
      '<p>Dacă ai întrebări sau vrei să schimbi ceva, răspunde la acest e-mail.</p>' +
      '<p>Cu drag,<br>Ana, de la Ce facem cu copiii?</p>';
  } else if (tip === 'colaborare') {
    subiect = 'Am primit mesajul tău - Ce facem cu copiii?';
    continut =
      '<p>Bună, ' + esc(c['Persoană de contact'] || '') + '!</p>' +
      '<p>Mulțumim pentru interesul de a colabora cu Ce facem cu copiii?. Am primit detaliile despre <strong>' + esc(c['Afacere / brand']) + '</strong> și revenim în curând, ca să ne cunoaștem și să-ți propunem variante potrivite.</p>' +
      '<p>Cu drag,<br>Ana, de la Ce facem cu copiii?</p>';
  } else {
    return;
  }
  MailApp.sendEmail({ to: catre, subject: subiect, htmlBody: continut, name: NUME_EXPEDITOR });
}

function esc(s) {
  return String(s || '').replace(/[&<>"]/g, function (ch) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch];
  });
}

/** Rulează o dată din editor: aprobi permisiunile și primești un mail de probă. */
function testMail() {
  trimite_confirmare('abonare', { 'E-mail': Session.getActiveUser().getEmail(), 'Vrea oferte de la parteneri': 'Da' });
}

/* ---------- 2. din „Propuneri evenimente” în „Evenimente” ---------- */

const STATUS_MUTA = 'Pregătit pentru calendar';
const STATUS_MUTAT = 'Mutat în calendar';
const ZILE = ['Duminică', 'Luni', 'Marți', 'Miercuri', 'Joi', 'Vineri', 'Sâmbătă'];

function onEdit(e) {
  try {
    const fila = e.range.getSheet();
    if (fila.getName() !== FILE.eveniment.nume) return;
    const colStatus = FILE.eveniment.coloane.length; // ultima coloană
    if (e.range.getColumn() !== colStatus || e.range.getRow() < 2) return;
    if (String(e.value || '') !== STATUS_MUTA) return;
    muta_in_calendar(fila, e.range.getRow());
  } catch (err) {
    e.source.toast('Nu am putut muta evenimentul: ' + err, 'Ce facem cu copiii?', 10);
  }
}

function muta_in_calendar(filaProp, rand) {
  const ss = filaProp.getParent();
  const tz = ss.getSpreadsheetTimeZone();
  const nr = FILE.eveniment.coloane.length;
  const v = filaProp.getRange(rand, 1, 1, nr).getValues()[0];
  const p = {};
  FILE.eveniment.coloane.forEach(function (c, i) { p[c[0]] = v[i]; });

  const data = ca_data(p['Data'], tz);
  if (!data) throw new Error('data nu e completată corect');
  const ora = ca_ora(p['Ora'], tz);
  const dataTxt = Utilities.formatDate(data, tz, 'dd.MM.yyyy');
  const acces = String(p['Acces'] || '');
  const bilet = acces === 'Gratuit' ? 'Gratuit' : acces === 'Cu bilet' ? 'Cu bilet' : acces === 'Cu înscriere' ? 'Cu înscriere' : '';
  const contact = [p['E-mail'], p['Telefon']].filter(String).join(', ');
  const primit = p['Primit la'] instanceof Date ? Utilities.formatDate(p['Primit la'], tz, 'dd.MM.yyyy HH:mm') : String(p['Primit la'] || '');

  const ev = ss.getSheetByName('Evenimente');
  // rândul nou: coloanele A-AC din „Evenimente”
  const nou = [
    '', 'De verificat', 'Da', "'" + dataTxt, ZILE[data.getDay()], ora ? "'" + ora : '', '',
    p['Organizator'], p['Nume eveniment'], '', p['Vârstă minimă'], p['Vârstă maximă'],
    p['Locație'], p['Adresă'], '', bilet, '', p['Link bilete / înscriere'], '',
    acces === 'Cu înscriere' || acces === 'Cu bilet' ? 'Da' : '', '', '', p['Descriere'], '',
    contact, 'Formular site', 'Organizator', Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy'),
    'Din propunerea primită pe ' + primit + '. De completat: categoria.'
  ];

  // îl punem la locul lui, după zi și oră
  const cand = new Date(data.getTime());
  if (ora) { const hm = ora.split(':'); cand.setHours(+hm[0], +hm[1]); }
  const prim = 3, ultim = Math.max(ev.getLastRow(), prim);
  const ordini = ev.getRange(prim, 30, ultim - prim + 1, 1).getValues();
  let tinta = ultim + 1;
  for (let i = 0; i < ordini.length; i++) {
    const o = ordini[i][0];
    if (o instanceof Date && o.getTime() > cand.getTime()) { tinta = prim + i; break; }
  }
  if (tinta <= ultim) ev.insertRowBefore(tinta);
  ev.getRange(tinta, 1, 1, nou.length).setValues([nou]);
  // formula din coloana „Ordine”, copiată de pe rândul vecin
  const vecin = tinta > prim ? tinta - 1 : tinta + 1;
  const f = ev.getRange(vecin, 30).getFormulaR1C1();
  if (f) ev.getRange(tinta, 30).setFormulaR1C1(f);

  filaProp.getRange(rand, nr).setValue(STATUS_MUTAT);
  ss.toast('Am mutat „' + p['Nume eveniment'] + '” în Evenimente, rândul ' + tinta + ', cu status „De verificat”.', 'Ce facem cu copiii?', 8);
}

function ca_data(x, tz) {
  if (x instanceof Date) return x;
  const s = String(x || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function ca_ora(x, tz) {
  if (x instanceof Date) return Utilities.formatDate(x, tz, 'HH:mm');
  const m = String(x || '').match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
