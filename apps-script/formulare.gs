/**
 * Ce facem cu copiii? - scriptul din sheet-ul principal (centralizatorul de evenimente).
 *
 * 1. Primește formularele de pe site:
 *    - propunerile de evenimente -> fila „Propuneri evenimente” din acest sheet
 *    - abonările și colaborările -> sheet-ul separat „Contacte”
 *    - la fiecare formular, o notificare pentru echipă (fila „Notificări” din sheet-ul cu texte)
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
const EMAIL_EXPEDITOR = 'hello@cefacemcucopiii.ro'; // expeditor verificat în Brevo
const SITE = 'https://cefacemcucopiii.ro';
const INSTAGRAM = 'https://www.instagram.com/cefacemcucopiii/';
const TEXTE_ID = '1bMT3_7UeHhK_FRl1pjzKYrAsXWvAXCNDzxvyqv5g6Yw'; // „Ce facem cu copiii? - Texte e-mailuri automate”
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
    try { notifica_echipa(date.tip, campuri); } catch (errNotif) { console.warn('Notificare: ' + errNotif); }
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

/* Textele mailurilor stau în fila „Texte e-mailuri” din sheet-ul separat „Ce facem cu copiii? - Texte e-mailuri automate” (le editezi acolo, fără cod).
   Coloane: Mail | Când pleacă | Subiect | Text. În text poți folosi: {persoana}, {brand}, {eveniment}, {pagina}, {site}, {instagram}.
   Un rând gol între paragrafe = paragraf nou. Dacă fila sau rândul lipsește, se folosește textul de rezervă de mai jos. */
const FILA_TEXTE = 'Texte e-mailuri';
const TEXTE_REZERVA = {
  'abonare': ['Bine ai venit la Ce facem cu copiii?', 'Ceau!\n\nMulțumim că te-ai abonat. De acum, în fiecare săptămână îți trimitem programul: spectacole, ateliere, concerte și ieșiri pentru copii din Timișoara.\n\n{oferte}\n\nPână la primul e-mail, tot calendarul e pe {site}, iar noutățile zilnice pe {instagram}.\n\nCu drag,\nAna, de la Ce facem cu copiii?\n\nTe poți dezabona oricând, răspunzând la acest e-mail cu „dezabonare”.'],
  'abonare - oferte': ['', 'Ai ales să primești și reduceri și oferte speciale pentru familii, de la organizatorii și brandurile cu care colaborăm. Le trimitem doar când merită.'],
  'eveniment': ['Am primit evenimentul tău: {eveniment}', 'Bună!\n\nMulțumim că ne-ai trimis {eveniment}. Îl verificăm și îl adăugăm în calendarul de pe {site}.\n\nDacă ai întrebări sau vrei să schimbi ceva, răspunde la acest e-mail.\n\nCu drag,\nAna, de la Ce facem cu copiii?'],
  'colaborare': ['Am primit mesajul tău - Ce facem cu copiii?', 'Bună, {persoana}!\n\nMulțumim pentru interesul de a colabora cu Ce facem cu copiii?. Am primit detaliile despre {brand} și revenim în curând, ca să ne cunoaștem și să-ți propunem variante potrivite.\n\nCu drag,\nAna, de la Ce facem cu copiii?'],
  'pagina live': ['Pagina {pagina} e live pe Ce facem cu copiii?', 'Bună, {persoana}!\n\nVești bune: pagina {pagina} e acum live în Comunitatea CFCC. O găsiți aici: {linkpagina}\n\nDe acum numărăm, în fiecare lună, cum interacționează părinții cu voi pe site: vizitele pe pagina voastră din Comunitate și ce fac pe evenimentele voastre din calendar (deschid detaliile, apasă pe bilete sau înscriere, le salvează, le pun în calendar, le trimit mai departe). Sunt cifre totale și anonime, nu urmărim persoane. La final de lună primiți raportul, așa cum am stabilit.\n\nVreți să schimbați ceva pe pagina de prezentare? Scrieți-ne la hello@cefacemcucopiii.ro sau răspundeți la acest e-mail și facem modificarea în cel mult 48 de ore.\n\nEvenimentele noi le puteți adăuga singuri, oricând, din meniul Colaborări → Adaugă un eveniment: {linkeveniment}. Le verificăm și apar în calendar.\n\nMulțumim că sunteți parte din Comunitatea CFCC!\n\nCu drag,\nAna, de la Ce facem cu copiii?'],
  'pagina': ['Am primit informațiile pentru pagina voastră - Ce facem cu copiii?', 'Bună, {persoana}!\n\nMulțumim! Am primit informațiile pentru pagina {pagina} din Comunitatea CFCC. Un om din echipa noastră le verifică și vă contactează dacă mai avem nevoie de ceva. Când pagina e gata, v-o trimitem spre aprobare și o publicăm doar după ce ne dați ok.\n\nDacă vreți să schimbați ceva între timp, răspundeți la acest e-mail.\n\nCu drag,\nAna, de la Ce facem cu copiii?']
};
const CAND_PLEACA = {
  'abonare': 'Cineva se abonează la newsletter pe site',
  'abonare - oferte': 'Paragraf pus în mailul de abonare doar dacă a bifat ofertele (înlocuiește {oferte}); subiectul nu contează',
  'eveniment': 'Cineva trimite un eveniment (Colaborări → Adaugă eveniment)',
  'colaborare': 'Cineva completează „Vreau să colaborăm”',
  'pagina': 'Un partener completează formularul pentru pagina lui (#formular-partener)',
  'pagina live': 'În Onboarding, statusul partenerului devine „Publicat” (pagina apare pe site). Pleacă o singură dată, la e-mailul persoanei de contact'
};

function texte_mailuri() {
  const ss = SpreadsheetApp.openById(TEXTE_ID);
  let f = ss.getSheetByName(FILA_TEXTE);
  if (!f) {
    f = ss.getSheets()[0];
    f.setName(FILA_TEXTE);
    const rows = [['Mail', 'Când pleacă', 'Subiect', 'Text']].concat(Object.keys(TEXTE_REZERVA).map(function (k) {
      return [k, CAND_PLEACA[k], TEXTE_REZERVA[k][0], TEXTE_REZERVA[k][1]];
    }));
    f.getRange(1, 1, rows.length, 4).setValues(rows);
    f.setFrozenRows(1);
    f.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1F4E5F').setFontColor('#ffffff');
    f.setColumnWidth(1, 140); f.setColumnWidth(2, 260); f.setColumnWidth(3, 320); f.setColumnWidth(4, 620);
    f.getRange(2, 1, rows.length - 1, 4).setWrap(true).setVerticalAlignment('top');
    f.getRange('F1').setValue('Poți folosi în Subiect și Text: {persoana}, {brand}, {eveniment}, {pagina}, {site}, {instagram}. Un rând gol = paragraf nou. Nu schimba coloana „Mail”.');
  }
  const t = {};
  f.getRange(2, 1, Math.max(f.getLastRow() - 1, 1), 4).getValues().forEach(function (r) {
    if (r[0]) t[String(r[0]).trim()] = [String(r[2] || ''), String(r[3] || '')];
  });
  return t;
}

function trimite_confirmare(tip, c) {
  const catre = String(c['E-mail'] || '').trim();
  if (!/^\S+@\S+\.\S+$/.test(catre)) return;
  if (!TEXTE_REZERVA[tip]) return;
  let t = {};
  try { t = texte_mailuri(); } catch (err) { console.warn('Texte: ' + err); }
  const ia = function (k) { return t[k] && t[k][1].trim() ? t[k] : TEXTE_REZERVA[k]; };
  const val = {
    persoana: c['Persoană de contact'] || '', brand: c['Afacere / brand'] || '',
    eveniment: c['Nume eveniment'] || '', pagina: c['Nume'] || ''
  };
  const link_pagina = SITE + '/#comunitate/' + (c._id || '');
  const link_eveniment = SITE + '/#adauga-eveniment';
  const subiect = ia(tip)[0].replace(/\{(\w+)\}/g, function (m, k) { return k in val ? val[k] : m; });
  let text = ia(tip)[1];
  text = text.replace('{oferte}', tip === 'abonare' && c['Vrea oferte de la parteneri'] === 'Da' ? ia('abonare - oferte')[1] : '');
  const continut = text.split(/\n\s*\n/).map(function (p) { return p.trim(); }).filter(String).map(function (p) {
    let h = esc(p).replace(/\n/g, '<br>');
    h = h.replace(/\{linkpagina\}/g, '<a href="' + link_pagina + '">' + esc(link_pagina.replace(/^https:\/\//, '')) + '</a>')
      .replace(/\{linkeveniment\}/g, '<a href="' + link_eveniment + '">' + esc(link_eveniment.replace(/^https:\/\//, '')) + '</a>')
      .replace(/\{site\}/g, '<a href="' + SITE + '">cefacemcucopiii.ro</a>')
      .replace(/\{instagram\}/g, '<a href="' + INSTAGRAM + '">Instagram</a>')
      .replace(/\{(persoana|brand|eveniment|pagina)\}/g, function (m, k) { return k === 'persoana' ? esc(val[k]) : '<strong>' + esc(val[k]) + '</strong>'; });
    return '<p>' + h + '</p>';
  }).join('');
  trimite_email(catre, subiect, continut);
}

/* ---------- notificări pentru echipă ---------- */

/* Fila „Notificări” din sheet-ul cu texte: Formular | Trimite (Da/Nu) | Către (una sau mai multe adrese, cu virgulă).
   Dacă fila lipsește, se creează singură, cu notificările pornite spre cefacemcucopiii@gmail.com. */
const FILA_NOTIFICARI = 'Notificări';
const NOTIF = {
  'abonare': ['Abonat nou', function (c) { return c['E-mail'] || ''; }, ['E-mail', 'Vrea oferte de la parteneri'], 'Contacte → Abonați', CONTACTE_ID],
  'eveniment': ['Eveniment propus', function (c) { return (c['Nume eveniment'] || '') + (c['Organizator'] ? ' (' + c['Organizator'] + ')' : ''); },
    ['Nume eveniment', 'Organizator', 'Data', 'Ora', 'Locație', 'Vârstă minimă', 'Vârstă maximă', 'Acces', 'E-mail', 'Telefon'], 'centralizator → Propuneri evenimente', ''],
  'colaborare': ['Cerere de colaborare', function (c) { return c['Afacere / brand'] || ''; },
    ['Afacere / brand', 'Persoană de contact', 'E-mail', 'Telefon', 'Instagram', 'Despre ei', 'Ce au nevoie'], 'Contacte → Colaborări', CONTACTE_ID],
  'pagina': ['Pagină de partener de verificat', function (c) { return c['Nume'] || ''; },
    ['Nume', 'Cum se descriu', 'Persoană de contact', 'Telefon contact', 'E-mail', 'Denumire legală'], 'CFCC - Onboarding parteneri', ONBOARDING_ID]
};

function setari_notificari() {
  const ss = SpreadsheetApp.openById(TEXTE_ID);
  let f = ss.getSheetByName(FILA_NOTIFICARI);
  if (!f) {
    f = ss.insertSheet(FILA_NOTIFICARI);
    const rows = [['Formular', 'Trimite notificare (Da/Nu)', 'Către', 'Ce primești']].concat([
      ['abonare', 'Da', 'cefacemcucopiii@gmail.com', 'Un mail „[CFCC] Abonat nou” la fiecare abonare la newsletter'],
      ['eveniment', 'Da', 'cefacemcucopiii@gmail.com', 'Un mail „[CFCC] Eveniment propus” când cineva trimite un eveniment'],
      ['colaborare', 'Da', 'cefacemcucopiii@gmail.com', 'Un mail „[CFCC] Cerere de colaborare” la fiecare cerere'],
      ['pagina', 'Da', 'cefacemcucopiii@gmail.com', 'Un mail „[CFCC] Pagină de partener de verificat” când un partener trimite formularul']
    ]);
    f.getRange(1, 1, rows.length, 4).setValues(rows);
    f.setFrozenRows(1);
    f.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1F4E5F').setFontColor('#ffffff');
    f.setColumnWidth(1, 120); f.setColumnWidth(2, 190); f.setColumnWidth(3, 300); f.setColumnWidth(4, 520);
    f.getRange(2, 2, rows.length - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Da', 'Nu'], true).build());
  }
  const s = {};
  f.getRange(2, 1, Math.max(f.getLastRow() - 1, 1), 3).getValues().forEach(function (r) {
    if (r[0]) s[String(r[0]).trim()] = { da: /^da/i.test(String(r[1])), catre: String(r[2] || '').trim() };
  });
  return s;
}

function notifica_echipa(tip, c) {
  const n = NOTIF[tip];
  if (!n) return;
  const s = setari_notificari()[tip];
  if (!s || !s.da || !s.catre) return;
  const subiect = '[CFCC] ' + n[0] + (n[1](c) ? ': ' + n[1](c) : '');
  const rows = n[2].filter(function (k) { return c[k]; }).map(function (k) {
    return '<tr><td style="padding:4px 12px 4px 0;color:#5b7079;vertical-align:top">' + esc(k) + '</td><td style="padding:4px 0">' + esc(c[k]).replace(/\n/g, '<br>') + '</td></tr>';
  }).join('');
  const link = n[4] ? 'https://docs.google.com/spreadsheets/d/' + n[4] + '/edit' : SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const html = '<p><strong>' + esc(n[0]) + '</strong>, primit acum pe cefacemcucopiii.ro.</p><table>' + rows + '</table>' +
    '<p>Toate detaliile sunt în <a href="' + link + '">' + esc(n[3]) + '</a>.</p>' +
    '<p style="color:#888;font-size:12px">Notificare automată. O oprești din sheet-ul „Ce facem cu copiii? - Texte e-mailuri automate”, fila Notificări.</p>';
  s.catre.split(/[,;\s]+/).filter(function (a) { return /^\S+@\S+\.\S+$/.test(a); }).forEach(function (a) { trimite_email(a, subiect, html); });
}

/* Trimite prin Brevo, de pe hello@cefacemcucopiii.ro.
   Cheia stă în Project Settings → Script Properties → BREVO_API_KEY (nu în cod).
   Dacă lipsește cheia sau Brevo dă eroare, trimite ca înainte, din Gmail. */
function trimite_email(catre, subiect, html) {
  const cheie = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
  if (cheie) {
    try {
      const r = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'api-key': cheie, accept: 'application/json' },
        payload: JSON.stringify({
          sender: { name: NUME_EXPEDITOR, email: EMAIL_EXPEDITOR },
          replyTo: { name: NUME_EXPEDITOR, email: EMAIL_EXPEDITOR },
          to: [{ email: catre }], subject: subiect, htmlContent: html
        })
      });
      if (r.getResponseCode() < 300) return;
      console.warn('Brevo ' + r.getResponseCode() + ': ' + r.getContentText());
    } catch (err) { console.warn('Brevo: ' + err); }
  }
  MailApp.sendEmail({ to: catre, subject: subiect, htmlBody: html, name: NUME_EXPEDITOR, replyTo: EMAIL_EXPEDITOR });
}

function esc(s) {
  return String(s || '').replace(/[&<>"]/g, function (ch) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch];
  });
}

/** Trimite pe adresa ta câte un mail de probă din fiecare tip (cu date de test). */
function testToateMailurile() {
  const eu = Session.getActiveUser().getEmail();
  trimite_confirmare('abonare', { 'E-mail': eu, 'Vrea oferte de la parteneri': 'Da' });
  trimite_confirmare('eveniment', { 'E-mail': eu, 'Nume eveniment': 'Atelier de test pentru copii' });
  trimite_confirmare('colaborare', { 'E-mail': eu, 'Persoană de contact': 'Ana', 'Afacere / brand': 'Brand de test' });
  trimite_confirmare('pagina', { 'E-mail': eu, 'Persoană de contact': 'Ana', 'Nume': 'Partener de test' });
}

/** Rulează o dată din editor: aprobi permisiunile și primești un mail de probă. */
function testMail() {
  testToateMailurile();
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
  const val = String(e.value || '');
  // Trimis spre aprobare / Publicat în Onboarding -> „Pe site” în Comunitatea CFCC (În aprobare / Da)
  if (/aprobare/i.test(val) || /^publicat/i.test(val)) {
    try {
      const pe = /aprobare/i.test(val) ? 'În aprobare' : 'Da';
      const id = slug_id(fila.getRange(e.range.getRow(), 3).getDisplayValue());
      const par = SpreadsheetApp.openById(COMUNITATE_ID).getSheetByName('Parteneri');
      const ids = par.getRange(1, 1, par.getLastRow(), 1).getDisplayValues().map(function (r) { return r[0]; });
      const i = ids.indexOf(id);
      if (i < 1) { e.source.toast('Nu am găsit pagina „' + id + '” în Comunitatea CFCC.', 'Ce facem cu copiii?', 10); return; }
      par.getRange(i + 1, 5).setValue(pe);
      let mesaj = 'În Comunitatea CFCC, „' + id + '” are acum Pe site = ' + pe + (pe === 'Da' ? ' (apare pe site în câteva minute).' : '.');
      if (pe === 'Da') mesaj += ' ' + mail_pagina_live(fila, e.range.getRow(), id);
      e.source.toast(mesaj, 'Ce facem cu copiii?', 12);
    } catch (err) { e.source.toast('Nu am putut actualiza Comunitatea CFCC: ' + err.message, 'Ce facem cu copiii?', 15); }
    return;
  }
  if (val !== STATUS_CREEAZA) return;
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

/** Trimite partenerului mailul „pagina e live”, o singură dată (se notează în coloana Note). */
function mail_pagina_live(fila, rand, id) {
  const nr = FILE.pagina.coloane.length;
  const v = fila.getRange(rand, 1, 1, nr).getDisplayValues()[0];
  const p = {};
  FILE.pagina.coloane.forEach(function (c, i) { p[c[0]] = String(v[i] || '').trim(); });
  const colNote = FILE.pagina.coloane.map(function (c) { return c[0]; }).indexOf('Note') + 1;
  if (/mail „pagina e live” trimis/i.test(p['Note'])) return 'Mailul „pagina e live” fusese deja trimis, nu l-am trimis din nou.';
  if (!/^\S+@\S+\.\S+$/.test(p['E-mail'])) return 'Nu am trimis mailul „pagina e live”: lipsește e-mailul persoanei de contact.';
  trimite_confirmare('pagina live', { 'E-mail': p['E-mail'], 'Persoană de contact': p['Persoană de contact'], 'Nume': p['Nume'], _id: id });
  const azi = Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy HH:mm');
  fila.getRange(rand, colNote).setValue((p['Note'] ? p['Note'] + '\n' : '') + azi + ' - mail „pagina e live” trimis la ' + p['E-mail']);
  return 'Am trimis mailul „pagina e live” la ' + p['E-mail'] + '.';
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

  // dacă pagina a mai existat și a fost ștearsă, scoatem activitățile și pozele vechi rămase cu același ID
  sterge_ramase(com.getSheetByName('Activități'), id);
  sterge_ramase(com.getSheetByName('Galerie'), id);

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

/** Șterge rândurile care au în coloana A ID-ul dat (resturi de la o pagină ștearsă). Articolele nu se ating. */
function sterge_ramase(fila, id) {
  if (!fila || fila.getLastRow() < 2) return;
  const ids = fila.getRange(2, 1, fila.getLastRow() - 1, 1).getDisplayValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]).trim() === id) fila.deleteRow(i + 2);
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
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'partener';
}

function link_social(x, retea) {
  x = String(x || '').trim();
  if (!x || /^https?:\/\//i.test(x)) return x;
  const nume = x.replace(/^@/, '').replace(/^(www\.)?(instagram|facebook)\.com\//i, '').replace(/\/$/, '');
  return 'https://www.' + retea + '.com/' + nume + '/';
}

function text(x) { return x ? "'" + x : ''; }
