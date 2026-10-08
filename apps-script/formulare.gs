/**
 * Ce facem cu copiii? - scriptul din sheet-ul principal (centralizatorul de evenimente).
 *
 * 1. Primește formularele de pe site:
 *    - propunerile de evenimente -> fila „Propuneri evenimente” din acest sheet
 *    - abonările și colaborările -> sheet-ul separat „Contacte”
 *    - formularul pentru parteneri (cefacemcucopiii.ro/#formular-partener) -> fila „Pagini parteneri - de verificat”
 *      din sheet-ul separat „CFCC - Onboarding parteneri”, cu roșu, ca să le verifice un om; pozele lor se salvează în Drive, în folderul „CFCC - poze parteneri”
 *    și trimite un e-mail de confirmare celui care a completat.
 * 2. Când schimbi statusul unei propuneri în „Pregătit pentru calendar”,
 *    o mută singur în fila „Evenimente”, la ziua și ora ei, cu status „De verificat”.
 *    După ce o treci pe „Confirmat”, apare pe site.
 * 3. În „CFCC - Onboarding parteneri”, când un rând primește statusul „Verificat - creează pagina”,
 *    îl copiază în „Comunitatea CFCC” ca ciornă (Pe site = Nu). Pornește o dată cu porneste_onboarding().
 */

const NUME_EXPEDITOR = 'Ce facem cu copiii?';
const SITE = 'https://cefacemcucopiii.ro';
const INSTAGRAM = 'https://www.instagram.com/cefacemcucopiii/';
const CONTACTE_ID = '1ydRGLLcpV0bGP07bckIoAZ4k2Z6llGmozHPJTVbSF38';
const ONBOARDING_ID = '1wyNnAA3GNJis2cofoQixTLizss0zU-90uBpFaaQIir0'; // „CFCC - Onboarding parteneri”

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
  },
  pagina: {
    nume: 'Pagini parteneri - de verificat', fisier: 'onboarding', status: 'De verificat', rosu: true, lungime: 5000,
    coloane: [['Primit la'], ['Status'], ['Nume'], ['Cum se descriu', 'Ce sunt'], ['Descriere'], ['Vârste'], ['Adresă'],
      ['Telefon public'], ['E-mail public'], ['Instagram'], ['Facebook'], ['Website'], ['Activități'], ['Nume în calendar'],
      ['Logo'], ['Poza principală'], ['Alte poze'], ['Folder poze'], ['Persoană de contact'], ['Telefon contact'], ['E-mail'],
      ['Denumire legală'], ['CUI'], ['Nr. înregistrare'], ['Sediu social'], ['Observații'], ['Acord publicare'], ['Note']]
  }
};

const FOLDER_POZE = 'CFCC - poze parteneri';

/* ---------- 1. formularele de pe site ---------- */

function doPost(e) {
  try {
    const date = JSON.parse(e.postData.contents);
    const tip = FILE[date.tip];
    if (!tip) return raspuns({ ok: false, eroare: 'tip necunoscut' });
    const campuri = date.campuri || {};
    if (date.tip === 'pagina') salveaza_poze(campuri);

    const fila = ia_fila(tip);
    const rand = tip.coloane.map(function (c) {
      const titlu = c[0], cheie = c[1] || c[0];
      if (titlu === 'Primit la') return new Date();
      if (titlu === 'Status') return tip.status || 'Nou';
      const v = campuri[cheie];
      if (v === undefined || v === null || v === '') return '';
      // apostroful păstrează textul exact cum a fost scris (telefoane cu 0 în față, date, ore)
      return "'" + String(v).slice(0, tip.lungime || 2000);
    });
    fila.appendRow(rand);
    if (tip.rosu) {
      fila.getRange(fila.getLastRow(), 1, 1, tip.coloane.length)
        .setBackground('#f4c7c3').setFontColor('#9c0006').setWrap(true).setVerticalAlignment('top');
    }

    try { trimite_confirmare(date.tip, campuri); } catch (errMail) { /* rândul e salvat oricum */ }
    return raspuns({ ok: true });
  } catch (err) {
    return raspuns({ ok: false, eroare: String(err) });
  }
}

function ia_fila(tip) {
  const ss = tip.fisier === 'contacte' ? SpreadsheetApp.openById(CONTACTE_ID)
    : tip.fisier === 'onboarding' ? SpreadsheetApp.openById(ONBOARDING_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let fila = ss.getSheetByName(tip.nume);
  if (!fila) {
    fila = ss.insertSheet(tip.nume);
    fila.appendRow(tip.coloane.map(function (c) { return c[0]; }));
    fila.setFrozenRows(1);
    fila.getRange(1, 1, 1, tip.coloane.length).setFontWeight('bold');
    if (tip.rosu) fila.getRange(1, 1, 1, tip.coloane.length).setBackground('#cc0000').setFontColor('#ffffff');
  }
  return fila;
}

/** Salvează pozele trimise de partener într-un folder al lui și pune linkurile în câmpuri. */
function salveaza_poze(c) {
  const poze = c._poze || {};
  delete c._poze;
  const toate = [].concat(poze.logo || [], poze.cover || [], poze.gal || []);
  if (!toate.length) return;
  const radacina = folder_sau_nou(DriveApp.getRootFolder(), FOLDER_POZE);
  const nume = String(c['Nume'] || 'Partener').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 80);
  const folder = radacina.createFolder(nume + ' - ' + Utilities.formatDate(new Date(), 'Europe/Bucharest', 'yyyy-MM-dd HH:mm'));
  const salveaza = function (lista, prefix) {
    return (lista || []).slice(0, 10).map(function (p, i) {
      const tipImg = /^image\/(png|jpeg)$/.test(p.tip) ? p.tip : 'image/jpeg';
      const blob = Utilities.newBlob(Utilities.base64Decode(p.data), tipImg, prefix + (i + 1) + '-' + String(p.nume || 'poza').slice(0, 60));
      return folder.createFile(blob).getUrl();
    }).join('\n');
  };
  c['Logo'] = salveaza(poze.logo, 'logo-');
  c['Poza principală'] = salveaza(poze.cover, 'principala-');
  c['Alte poze'] = salveaza(poze.gal, 'poza-');
  c['Folder poze'] = folder.getUrl();
}

function folder_sau_nou(parinte, nume) {
  const it = parinte.getFoldersByName(nume);
  return it.hasNext() ? it.next() : parinte.createFolder(nume);
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
  } else if (tip === 'pagina') {
    subiect = 'Am primit informațiile pentru pagina voastră - Ce facem cu copiii?';
    continut =
      '<p>Bună, ' + esc(c['Persoană de contact'] || '') + '!</p>' +
      '<p>Mulțumim! Am primit informațiile pentru pagina <strong>' + esc(c['Nume']) + '</strong> din Comunitatea CFCC. Un om din echipa noastră le verifică și vă contactează dacă mai avem nevoie de ceva. Când pagina e gata, v-o trimitem spre aprobare și o publicăm doar după ce ne dați ok.</p>' +
      '<p>Dacă vreți să schimbați ceva între timp, răspundeți la acest e-mail.</p>' +
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

/* ---------- 3. din „Onboarding parteneri” în „Comunitatea CFCC” ---------- */

const COMUNITATE_ID = '19VxFnrcRzF8sMIG0tUiQSexOE1I21g5tXm9CQw2n7hQ'; // „Comunitatea CFCC - pagini parteneri” (public)
const STATUS_CREEAZA = 'Verificat - creează pagina';
const STATUS_CREATA = 'Pagină creată (Pe site = Nu)';

/** Rulează o singură dată din editor: pornește copierea automată din sheet-ul de onboarding. */
function porneste_onboarding() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'la_editare_onboarding') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('la_editare_onboarding').forSpreadsheet(ONBOARDING_ID).onEdit().create();
}

function la_editare_onboarding(e) {
  const fila = e.range.getSheet();
  if (fila.getName() !== FILE.pagina.nume) return;
  if (e.range.getColumn() !== 2 || e.range.getRow() < 2) return; // coloana Status
  if (String(e.value || '') !== STATUS_CREEAZA) return;
  try {
    const id = creeaza_pagina(fila, e.range.getRow());
    e.range.setValue(STATUS_CREATA);
    fila.getRange(e.range.getRow(), 1, 1, FILE.pagina.coloane.length).setBackground('#d9ead3').setFontColor('#274e13');
    e.source.toast('Am creat pagina „' + id + '” în Comunitatea CFCC, ca ciornă (Pe site = Nu).', 'Ce facem cu copiii?', 10);
  } catch (err) {
    e.range.setValue('De verificat');
    e.source.toast('Nu am putut crea pagina: ' + err.message, 'Ce facem cu copiii?', 15);
  }
}

function creeaza_pagina(fila, rand) {
  const v = fila.getRange(rand, 1, 1, FILE.pagina.coloane.length).getDisplayValues()[0];
  const p = {};
  FILE.pagina.coloane.forEach(function (c, i) { p[c[0]] = String(v[i] || '').trim(); });
  if (!p['Nume']) throw new Error('lipsește numele');

  const com = SpreadsheetApp.openById(COMUNITATE_ID);
  const parteneri = com.getSheetByName('Parteneri');
  const id = slug_id(p['Nume']);
  const existente = parteneri.getRange(1, 1, Math.max(parteneri.getLastRow(), 1), 1).getDisplayValues().map(function (r) { return r[0]; });
  if (existente.indexOf(id) >= 0) throw new Error('există deja o pagină cu ID-ul „' + id + '”');

  const logo = poze_publice(p['Logo'])[0] || '';
  const cover = poze_publice(p['Poza principală'])[0] || '';
  const alte = poze_publice(p['Alte poze']);
  const nume_calendar = [p['Nume']].concat(p['Nume în calendar'].split(/\n+/)).map(function (x) { return x.trim(); }).filter(String).join('\n');

  parteneri.appendRow([id, p['Nume'], p['Cum se descriu'], 'Partener', 'Nu', p['Descriere'], p['Vârste'], p['Adresă'],
    text(p['Telefon public']), p['E-mail public'], link_social(p['Instagram'], 'instagram'), link_social(p['Facebook'], 'facebook'),
    p['Website'], logo, cover, nume_calendar]);

  const act = p['Activități'].split(/\n+/).filter(String).map(function (l) {
    const parti = l.split('|');
    return [id, parti[0].trim(), (parti.slice(1).join('|') || '').trim()];
  });
  if (act.length) {
    const fa = com.getSheetByName('Activități');
    fa.getRange(fa.getLastRow() + 1, 1, act.length, 3).setValues(act);
  }
  if (alte.length) {
    const fg = com.getSheetByName('Galerie');
    fg.getRange(fg.getLastRow() + 1, 1, alte.length, 3).setValues(alte.map(function (u) { return [id, u, '']; }));
  }
  return id;
}

/** Face pozele din Drive vizibile pentru oricine are linkul și întoarce linkuri care se pot afișa pe site. */
function poze_publice(celula) {
  return String(celula || '').split(/\s+/).filter(String).map(function (u) {
    const m = u.match(/\/d\/([\w-]{20,})/) || u.match(/[?&]id=([\w-]{20,})/);
    if (!m) return u;
    DriveApp.getFileById(m[1]).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://lh3.googleusercontent.com/d/' + m[1];
  });
}

function slug_id(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'partener';
}

function link_social(x, retea) {
  x = String(x || '').trim();
  if (!x || /^https?:\/\//i.test(x)) return x;
  const nume = x.replace(/^@/, '').replace(/^(www\.)?(instagram|facebook)\.com\//i, '').replace(/\/$/, '');
  return 'https://www.' + retea + '.com/' + nume + '/';
}

function text(x) { return x ? "'" + x : ''; }
