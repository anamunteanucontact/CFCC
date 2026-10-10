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
 * 2. Propuneri: Nou (roșu) → „În calendar - nepublicat” (portocaliu) o mută singur în „Evenimente”, la ziua și ora ei,
 *    cu status „De verificat”; „Publicat” (verde) o mută direct publicată. Statusul se sincronizează în ambele sensuri.
 * 3. În „CFCC - Onboarding parteneri”, când un rând primește statusul „Verificat - creează pagina”,
 *    face link de previzualizare, trimite mailul spre aprobare și, la „Publicat”, copiază pagina în „Comunitatea CFCC”. Pornește o dată cu porneste_onboarding().
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
      ['Vârstă minimă'], ['Vârstă maximă'], ['Acces'], ['Link bilete / înscriere'], ['Descriere'], ['E-mail'], ['Telefon'], ['Categorie'], ['Status'], ['Note']]
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
// Folderul fix pentru poze (în folderul principal CFCC). Nu se mai creează altul nou.
const FOLDER_POZE_ID = '1us3oNdxAI03kIt-6M3R6exsmeCBcBPRc';
const FOLDER_CFCC_ID = '1N3rD9icm9kMZJIVBwIw16FJyUPdDyxk9';

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
    const lacat = LockService.getScriptLock();
    lacat.waitLock(20000);
    let nrRand;
    try { fila.appendRow(rand); nrRand = fila.getLastRow(); } finally { lacat.releaseLock(); }
    if (tip.rosu) {
      fila.getRange(nrRand, 1, 1, tip.coloane.length)
        .setBackground('#f4c7c3').setFontColor('#9c0006').setWrap(true).setVerticalAlignment('top');
    }

    let rezultatMail;
    try { rezultatMail = trimite_confirmare(date.tip, campuri); } catch (errMail) { rezultatMail = 'eroare: ' + errMail; /* rândul e salvat oricum */ }
    try { noteaza_mail_confirmare(fila, nrRand, tip, date.tip, campuri, rezultatMail); } catch (errNota) { console.warn('Notă: ' + errNota); }
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
  const radacina = folder_poze();
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

/** Folderul de poze: cel fix, după ID. Doar dacă a fost șters, îl caută / creează în folderul principal CFCC. */
function folder_poze() {
  try { return DriveApp.getFolderById(FOLDER_POZE_ID); } catch (e) {}
  return folder_sau_nou(DriveApp.getFolderById(FOLDER_CFCC_ID), FOLDER_POZE);
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
  'pagina spre aprobare': ['Pagina {pagina} e gata - te rugăm să o verifici', 'Bună, {persoana}!\n\nPagina {pagina} din Comunitatea CFCC e gata. O poți vedea aici: {linkpagina}\n\nDeocamdată nu e publică: nu apare încă pe site, o vezi doar tu, prin acest link.\n\nTe rugăm să te uiți peste ea (texte, poze, date de contact, activități) și să ne răspunzi la acest e-mail:\n- dacă totul e în regulă, scrie-ne „Confirm” și o publicăm;\n- dacă vrei să schimbi ceva, spune-ne ce anume și o ajustăm înainte de publicare.\n\nPublicăm pagina doar după ce primim confirmarea ta.\n\nCu drag,\nAna, de la Ce facem cu copiii?'],
  'pagina live': ['Pagina {pagina} e live pe platforma noastră', 'Bună, {persoana}!\n\nVești bune: pagina {pagina} e acum live în Comunitatea CFCC.\nO găsești aici: {linkpagina}\n\nDe acum numărăm, în fiecare lună, cum interacționează părinții cu {pagina} pe site: vizitele pe pagina din Comunitate și ce fac pe evenimentele tale din calendar (deschid detaliile, apasă pe bilete sau înscriere, le salvează, le pun în calendar, le trimit mai departe).\nSunt cifre totale și anonime, nu urmărim persoane. La final de lună primești raportul, așa cum am stabilit.\n\nVrei să schimbi ceva pe pagina de prezentare?\nScrie-ne la hello@cefacemcucopiii.ro sau răspunde la acest e-mail și facem modificarea în cel mult 48 de ore.\n\nEvenimentele noi le poți adăuga tu, oricând, din meniul Colaborări → Adaugă un eveniment: {linkeveniment}.\nLe verificăm și apoi apar în calendar.\n\nMulțumim că ești parte din Comunitatea CFCC!\n\nCu drag,\nAna, de la Ce facem cu copiii?'],
  'pagina live - recomandări': ['Câteva sfaturi ca pagina {pagina} să fie văzută de cât mai mulți părinți', 'Bună, {persoana}!\n\nPrima zi pe platformă s-a încheiat și am început deja să strângem datele pentru raportul de la final de lună.\n\nCa pagina să lucreze cât mai mult pentru tine, câteva lucruri simple care ajută mult la vizibilitate:\n\n1. Pune linkul paginii în bio, pe Instagram și pe Facebook: {linkpagina}\n2. Când îți dăm tag într-o postare sau într-un story, dă share sau repost, și în story, și pe profil.\n3. Când îți trimitem o invitație de colaborare (collab) pe Instagram, accept-o: postarea apare și pe profilul tău și ajunge la ambele comunități.\n4. Urmărește-ne pe {instagram} și dă-ne tag când postezi ceva pentru copii, ca să putem distribui mai departe.\n5. Trimite-ne evenimentele noi din timp, din Colaborări → Adaugă un eveniment: {linkeveniment}.\nCu cât apar mai devreme în calendar, cu atât le văd mai mulți părinți.\n\nPentru orice întrebare, suntem la hello@cefacemcucopiii.ro.\n\nCu drag,\nAna, de la Ce facem cu copiii?'],
  'pagina': ['Gata, am primit informațiile pentru pagina ta!', 'Bună, {persoana}!\n\nMulțumim! Am primit informațiile pentru pagina {pagina} din Comunitatea CFCC.\n\nUn om (adevărat) din echipa noastră le verifică chiar acum și te contactăm cât mai repede, dacă mai avem nevoie de ceva.\n\nRămâne cum am stabilit: când pagina e gata, ți-o trimitem spre aprobare și o publicăm doar după ce ne dai ok.\n\nDacă vrei să schimbi ceva între timp, răspunde la acest e-mail.\n\nCu drag,\nAna, de la Ce facem cu copiii?']
};
TEXTE_REZERVA['eveniment publicat'] = ['{eveniment} e acum în calendarul Ce facem cu copiii?', 'Bună!\n\nVești bune: {eveniment} e acum în calendarul de pe {site}, locul unde părinții din Timișoara găsesc tot ce e de făcut cu copiii.\n\nLinkul direct spre eveniment: {link}\n\nCa să ajungă la cât mai mulți părinți, te rugăm:\n1. Pune linkul în postările și story-urile despre eveniment, pe Instagram și pe Facebook.\n2. Pune-l în bio sau în linkurile din profil, cât timp promovezi evenimentul.\n3. Trimite-l mai departe pe grupurile de WhatsApp ale părinților.\n4. Dă-ne tag (@cefacemcucopiii) când postezi, ca să putem distribui și noi.\n\nAi și alte evenimente pentru copii? Le poți adăuga oricând aici: {linkeveniment}\n\nCu drag,\nAna, de la Ce facem cu copiii?'];
const CAND_PLEACA = {
  'abonare': 'Cineva se abonează la newsletter pe site',
  'abonare - oferte': 'Paragraf pus în mailul de abonare doar dacă a bifat ofertele (înlocuiește {oferte}); subiectul nu contează',
  'eveniment': 'Cineva trimite un eveniment (Colaborări → Adaugă eveniment)',
  'colaborare': 'Cineva completează „Vreau să colaborăm”',
  'pagina': 'Un partener completează formularul pentru pagina lui (#formular-partener)',
  'pagina spre aprobare': 'În Onboarding, statusul devine „Trimis spre aprobare”. Trimite partenerului linkul de previzualizare și îi cere confirmarea prin reply. Pleacă o singură dată',
  'pagina live': 'În Onboarding, statusul partenerului devine „Publicat” (pagina apare pe site). Pleacă o singură dată, la e-mailul persoanei de contact',
  'eveniment publicat': 'Un eveniment trimis de organizator prin formular primește Status = Publicat în The Sheet (verificare la 15 minute, cam la 10-25 de minute după publicare, ca să fie deja pe site). Pleacă o singură dată, la e-mailul din Contact organizator. {link} = linkul evenimentului de pe site',
  'pagina live - recomandări': 'La 1-2 zile după mailul „pagina live” (verificare zilnică la 10:00). Pleacă o singură dată, doar dacă statusul e tot „Publicat”.'
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
  if (!/^\S+@\S+\.\S+$/.test(catre)) return '';
  if (!TEXTE_REZERVA[tip]) return '';
  let t = {};
  try { t = texte_mailuri(); } catch (err) { console.warn('Texte: ' + err); }
  const ia = function (k) { return t[k] && t[k][1].trim() ? t[k] : TEXTE_REZERVA[k]; };
  const val = {
    persoana: c['Persoană de contact'] || '', brand: c['Afacere / brand'] || '',
    eveniment: c['Nume eveniment'] || '', pagina: c['Nume'] || ''
  };
  const link_pagina = c._link || (SITE + '/#comunitate/' + (c._id || ''));
  const link_eveniment = SITE + '/#adauga-eveniment';
  const subiect = ia(tip)[0].replace(/\{(\w+)\}/g, function (m, k) { return k in val ? val[k] : m; });
  let text = ia(tip)[1];
  text = text.replace('{oferte}', tip === 'abonare' && c['Vrea oferte de la parteneri'] === 'Da' ? ia('abonare - oferte')[1] : '');
  const continut = text.split(/\n\s*\n/).map(function (p) { return p.trim(); }).filter(String).map(function (p) {
    let h = esc(p).replace(/\n/g, '<br>');
    h = h.replace(/\{linkpagina\}/g, '<a href="' + link_pagina + '">' + esc(link_pagina.replace(/^https:\/\//, '')) + '</a>')
      .replace(/\{linkeveniment\}/g, '<a href="' + link_eveniment + '">' + esc(link_eveniment.replace(/^https:\/\//, '')) + '</a>')
      .replace(/\{link\}/g, c._linkEv ? '<a href="' + c._linkEv + '">' + esc(c._linkEv.replace(/^https:\/\//, '')) + '</a>' : '')
      .replace(/\{site\}/g, '<a href="' + SITE + '">cefacemcucopiii.ro</a>')
      .replace(/\{instagram\}/g, '<a href="' + INSTAGRAM + '">Instagram</a>')
      .replace(/\{(persoana|brand|eveniment|pagina)\}/g, function (m, k) { return k === 'persoana' ? esc(val[k]) : '<strong>' + esc(val[k]) + '</strong>'; });
    return '<p>' + h + '</p>';
  }).join('');
  return trimite_email(catre, subiect, continut);
}

/* Ce mail de confirmare pleacă după fiecare formular (numele care apare în coloana Note). */
const ETICHETE_CONFIRMARE = { 'eveniment': 'am primit evenimentul', 'abonare': 'bine ai venit', 'colaborare': 'confirmare colaborare', 'pagina': 'confirmare formular' };

/** Notează în coloana Note a rândului nou dacă mailul de confirmare a plecat (cu data, ora și adresa). */
function noteaza_mail_confirmare(fila, nrRand, tip, cheieTip, c, rezultat) {
  const colNote = tip.coloane.map(function (x) { return x[0]; }).indexOf('Note') + 1;
  const eticheta = ETICHETE_CONFIRMARE[cheieTip];
  if (!colNote || !eticheta) return;
  const catre = String(c['E-mail'] || '').trim();
  const azi = Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy HH:mm');
  let nota;
  if (rezultat === 'Brevo' || rezultat === 'Gmail') nota = azi + ' - mail „' + eticheta + '” trimis la ' + catre;
  else if (!catre) nota = azi + ' - mail „' + eticheta + '” NU a plecat: lipsește adresa de e-mail';
  else nota = azi + ' - mail „' + eticheta + '” NU a plecat la ' + catre + (rezultat ? ' (' + String(rezultat).slice(0, 200) + ')' : '');
  const celula = fila.getRange(nrRand, colNote);
  const vechi = celula.getDisplayValue();
  celula.setValue((vechi ? vechi.replace(/\s+$/, '') + '\n\n' : '') + nota);
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
  const html = '<p><strong>' + esc(n[0]) + '</strong>' + (n[5] ? '. ' + esc(n[5]) : ', primit acum pe cefacemcucopiii.ro.') + '</p><table>' + rows + '</table>' +
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
      if (r.getResponseCode() < 300) return 'Brevo';
      console.warn('Brevo ' + r.getResponseCode() + ': ' + r.getContentText());
    } catch (err) { console.warn('Brevo: ' + err); }
  }
  MailApp.sendEmail({ to: catre, subject: subiect, htmlBody: html, name: NUME_EXPEDITOR, replyTo: EMAIL_EXPEDITOR });
  return 'Gmail';
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

/** Trimite pe adresa ta mailurile de după onboarding („pagina live” și „recomandări”), cu textele din sheet, pe exemplul ABCD Creativity. */
function testMailuriParteneri() {
  const eu = Session.getActiveUser().getEmail();
  const c = { 'E-mail': eu, 'Persoană de contact': 'Ana', 'Nume': 'ABCD Creativity', _id: 'abcd-creativity' };
  trimite_confirmare('pagina live', c);
  trimite_confirmare('pagina live - recomandări', c);
}

/** Rulează o dată din editor: aprobi permisiunile și primești un mail de probă. */
function testMail() {
  testToateMailurile();
}

/* ---------- 2. din „Propuneri evenimente” în „Evenimente” ---------- */

// Statusurile din „Propuneri evenimente”: Nou (roșu) → În calendar - nepublicat (portocaliu) → Publicat (verde); Respins (gri).
const STATUS_IN_CALENDAR = 'În calendar - nepublicat';
const STATUS_PUBLICAT = 'Publicat';
const NOTA_MUTAT = 'Mutat în Evenimente';
const ZILE = ['Duminică', 'Luni', 'Marți', 'Miercuri', 'Joi', 'Vineri', 'Sâmbătă'];

function onEdit(e) {
  try {
    const fila = e.range.getSheet();
    if (fila.getName() === 'Evenimente') { la_editare_evenimente(e); return; }
    if (fila.getName() !== FILE.eveniment.nume) return;
    const colStatus = prop_col('Status');
    if (e.range.getColumn() !== colStatus || e.range.getRow() < 2) return;
    const val = String(e.value || '');
    if (val !== STATUS_IN_CALENDAR && val !== STATUS_PUBLICAT) return;
    const celula = e.range;
    const colCat = FILE.eveniment.coloane.map(function (c) { return c[0]; }).indexOf('Categorie') + 1;
    if (val === STATUS_PUBLICAT && !String(fila.getRange(e.range.getRow(), colCat).getValue() || '').trim()) {
      celula.setValue(e.oldValue || 'Nou');
      e.source.toast('Alege întâi categoria (coloana Categorie), apoi pune „Publicat”.', 'Ce facem cu copiii?', 8);
      return;
    }
    if (String(celula.getNote() || '').indexOf(NOTA_MUTAT) === 0) {
      // e deja în Evenimente: doar sincronizăm statusul acolo
      seteaza_status_in_evenimente(fila, e.range.getRow(), val === STATUS_PUBLICAT ? 'Publicat' : 'De verificat');
      return;
    }
    muta_in_calendar(fila, e.range.getRow(), val === STATUS_PUBLICAT);
  } catch (err) {
    e.source.toast('Nu am putut muta evenimentul: ' + err, 'Ce facem cu copiii?', 10);
  }
}

/** Coloana (1, 2, ...) dintr-un titlu din „Propuneri evenimente”. */
function prop_col(titlu) { return FILE.eveniment.coloane.map(function (c) { return c[0]; }).indexOf(titlu) + 1; }

/** Adaugă un rând în coloana Note a propunerii (cu data și ora). */
function noteaza_propunere(filaProp, rand, text) {
  const c = filaProp.getRange(rand, prop_col('Note'));
  const azi = Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy HH:mm');
  const vechi = c.getDisplayValue();
  c.setValue((vechi ? vechi.replace(/\s+$/, '') + '\n\n' : '') + azi + ' - ' + text);
}

/** Găsește rândul propunerii după nume și dată (pentru notele venite din Evenimente). */
function rand_propunere(ss, nume, data) {
  const prop = ss.getSheetByName(FILE.eveniment.nume);
  if (!prop || prop.getLastRow() < 2) return 0;
  const tz = ss.getSpreadsheetTimeZone(), n = String(nume || '').trim().toLowerCase(), d = ca_data(data, tz);
  const rows = prop.getRange(2, 1, prop.getLastRow() - 1, 4).getValues();
  for (let i = 0; i < rows.length; i++) {
    const dd = ca_data(rows[i][3], tz);
    if (String(rows[i][1] || '').trim().toLowerCase() === n && d && dd && dd.toDateString() === d.toDateString()) return i + 2;
  }
  return 0;
}

function muta_in_calendar(filaProp, rand, publica) {
  const ss = filaProp.getParent();
  const tz = ss.getSpreadsheetTimeZone();
  const nr = FILE.eveniment.coloane.length;
  const v = filaProp.getRange(rand, 1, 1, nr).getValues()[0];
  const p = {};
  FILE.eveniment.coloane.forEach(function (c, i) { p[c[0]] = v[i]; });

  const ev = ss.getSheetByName('Evenimente');
  const data = ca_data(p['Data'], tz);
  if (!data) throw new Error('data nu e completată corect');
  const ora = ca_ora(p['Ora'], tz);
  const dataTxt = Utilities.formatDate(data, tz, 'dd.MM.yyyy');
  const acces = String(p['Acces'] || '');
  const bilet = acces === 'Gratuit' ? 'Gratuit' : acces === 'Cu bilet' ? 'Cu bilet' : acces === 'Cu înscriere' ? 'Cu înscriere' : '';
  const contact = [p['E-mail'], p['Telefon']].filter(String).join(', ');
  const primit = p['Primit la'] instanceof Date ? Utilities.formatDate(p['Primit la'], tz, 'dd.MM.yyyy HH:mm') : String(p['Primit la'] || '');

  // rândul nou: coloanele B-U din „Evenimente” (coloana A e formula cu linkul de pe site)
  const notaProp = 'Din propunerea primită pe ' + primit + '.' + (p['Categorie'] ? '' : ' De completat: categoria.');
  const nou = [
    publica ? 'Publicat' : 'De verificat', 'Nu', "'" + dataTxt, '', ora ? "'" + ora : '', '',
    p['Nume eveniment'], p['Organizator'], p['Locație'], p['Descriere'],
    p['Vârstă minimă'], p['Vârstă maximă'], p['Categorie'] || '', p['Adresă'] || adresa_cunoscuta(ev, p['Locație']),
    bilet, p['Link bilete / înscriere'], contact, notaProp,
    'Organizator', Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy')
  ];

  // îl punem la locul lui, după zi și oră
  const cand = new Date(data.getTime());
  if (ora) { const hm = ora.split(':'); cand.setHours(+hm[0], +hm[1]); }
  const prim = 3, ultim = Math.max(ev.getLastRow(), prim);
  const ordini = ev.getRange(prim, EV_COL.ordine, ultim - prim + 1, 1).getValues();
  let tinta = ultim + 1;
  for (let i = 0; i < ordini.length; i++) {
    const o = ordini[i][0];
    if (o instanceof Date && o.getTime() > cand.getTime()) { tinta = prim + i; break; }
  }
  if (tinta <= ultim) ev.insertRowBefore(tinta);
  ev.getRange(tinta, 2, 1, nou.length).setValues([nou]);
  // „Zi” și „Ordine” se calculează singure (ARRAYFORMULA în capul coloanelor, rândul 2), de aceea rămân goale aici.

  filaProp.getRange(rand, prop_col('Status')).setNote(NOTA_MUTAT + ' pe ' + Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm'));
  noteaza_propunere(filaProp, rand, 'mutat în Evenimente, rândul ' + tinta + (publica ? ', publicat pe site' : ', nepublicat încă'));
  ss.toast('Am mutat „' + p['Nume eveniment'] + '” în Evenimente, rândul ' + tinta + ', cu status „' + (publica ? 'Publicat' : 'De verificat') + '”.', 'Ce facem cu copiii?', 8);
}

/** Găsește în Evenimente rândul venit dintr-o propunere (același nume și aceeași dată) și îi pune statusul dat. */
function seteaza_status_in_evenimente(filaProp, rand, status) {
  const ss = filaProp.getParent(), tz = ss.getSpreadsheetTimeZone();
  const nr = FILE.eveniment.coloane.length;
  const v = filaProp.getRange(rand, 1, 1, nr).getValues()[0];
  const nume = String(v[1] || '').trim().toLowerCase(), d = ca_data(v[3], tz);
  const ev = ss.getSheetByName('Evenimente');
  if (ev.getLastRow() < 3 || !nume) return;
  const rows = ev.getRange(3, 1, ev.getLastRow() - 2, EV_COL.nume).getValues();
  for (let i = 0; i < rows.length; i++) {
    const dd = ca_data(rows[i][EV_COL.data - 1], tz);
    if (String(rows[i][EV_COL.nume - 1] || '').trim().toLowerCase() === nume && d && dd && dd.toDateString() === d.toDateString()) {
      ev.getRange(i + 3, EV_COL.status).setValue(status);
      ss.toast('Am pus „' + status + '” și în Evenimente, rândul ' + (i + 3) + '.', 'Ce facem cu copiii?', 6);
      return;
    }
  }
}

/** Invers: când schimbi Status-ul în Evenimente, propunerea din care a venit devine Publicat (verde) sau În calendar - nepublicat (portocaliu). */
function sincronizeaza_propunerea(ev, rand) {
  const ss = ev.getParent(), tz = ss.getSpreadsheetTimeZone();
  const r = ev.getRange(rand, 1, 1, EV_COL.nume).getValues()[0];
  const nume = String(r[EV_COL.nume - 1] || '').trim().toLowerCase(), d = ca_data(r[EV_COL.data - 1], tz);
  const st = String(r[EV_COL.status - 1] || '').trim();
  const prop = ss.getSheetByName(FILE.eveniment.nume);
  if (!prop || prop.getLastRow() < 2 || !nume) return;
  const nr = FILE.eveniment.coloane.length, cs = prop_col('Status');
  const rows = prop.getRange(2, 1, prop.getLastRow() - 1, nr).getValues();
  const note = prop.getRange(2, cs, prop.getLastRow() - 1, 1).getNotes();
  for (let i = 0; i < rows.length; i++) {
    if (String(note[i][0] || '').indexOf(NOTA_MUTAT) !== 0) continue;
    const dd = ca_data(rows[i][3], tz);
    if (String(rows[i][1] || '').trim().toLowerCase() === nume && d && dd && dd.toDateString() === d.toDateString()) {
      const nou = st === 'Publicat' ? STATUS_PUBLICAT : st === 'Anulat' ? 'Respins' : STATUS_IN_CALENDAR;
      if (String(rows[i][cs - 1] || '') !== nou) {
        prop.getRange(i + 2, cs).setValue(nou);
        noteaza_propunere(prop, i + 2, 'status schimbat din Evenimente: ' + nou);
      }
      return;
    }
  }
}

/* ---------- „Evenimente” (The Sheet): coloanele, după reorganizarea din 9 oct ---------- */
// A ID (link pe site, formulă) | B Status | C Recomandare | D Data | E Zi | F Ora start | G Ora sfârșit | H Nume eveniment
// I Organizator | J Locație | K Descriere | L Vârstă de la | M Vârstă până la | N Categorie | O Adresă | P Bilet
// Q Link bilete | R Contact organizator | S Alte detalii / note interne | T Adăugat de | U Data adăugării | V Ordine (formulă)
const EV_COL = { status: 2, data: 4, nume: 8, organizator: 9, locatie: 10, adresa: 15, adaugat: 21, ordine: 22 };

/** La editare în Evenimente: adresa se completează singură după locație, iar data adăugării se pune singură. */
function la_editare_evenimente(e) {
  const fila = e.range.getSheet();
  const r0 = e.range.getRow(), c0 = e.range.getColumn(), nr = e.range.getNumRows(), nc = e.range.getNumColumns();
  if (r0 + nr - 1 < 3) return;
  const tz = e.source.getSpreadsheetTimeZone();
  for (let r = Math.max(r0, 3); r < r0 + nr; r++) {
    // statusul s-a schimbat -> propunerea din care a venit își schimbă culoarea
    if (c0 <= EV_COL.status && EV_COL.status < c0 + nc) { try { sincronizeaza_propunerea(fila, r); } catch (errS) { console.warn(errS); } }
    // locația s-a schimbat și adresa e goală -> o luăm din intrările de dinainte sau din Organizatori
    if (c0 <= EV_COL.locatie && EV_COL.locatie < c0 + nc) {
      const loc = String(fila.getRange(r, EV_COL.locatie).getValue() || '').trim();
      const adr = fila.getRange(r, EV_COL.adresa);
      if (loc && !String(adr.getValue() || '').trim()) {
        const gasita = adresa_cunoscuta(fila, loc, r);
        if (gasita) { adr.setValue(gasita); e.source.toast('Am completat adresa pentru „' + loc + '”.', 'Ce facem cu copiii?', 4); }
      }
    }
    // rând nou (are nume sau dată) -> data adăugării, automat
    const ad = fila.getRange(r, EV_COL.adaugat);
    if (!String(ad.getValue() || '').trim()) {
      const are = String(fila.getRange(r, EV_COL.nume).getValue() || '').trim() || String(fila.getRange(r, EV_COL.data).getValue() || '').trim();
      if (are) ad.setValue("'" + Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy'));
    }
  }
}

/** Caută adresa unei locații: întâi în Evenimente (cea mai recentă intrare cu aceeași locație), apoi în Organizatori. */
function adresa_cunoscuta(fila, loc, randExclus) {
  const cheie = String(loc || '').trim().toLowerCase();
  if (!cheie) return '';
  const ultim = fila.getLastRow();
  if (ultim >= 3) {
    const v = fila.getRange(3, EV_COL.locatie, ultim - 2, EV_COL.adresa - EV_COL.locatie + 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (i + 3 === randExclus) continue;
      const a = String(v[i][EV_COL.adresa - EV_COL.locatie] || '').trim();
      if (a && String(v[i][0] || '').trim().toLowerCase() === cheie) return a;
    }
  }
  const org = fila.getParent().getSheetByName('Organizatori');
  if (org && org.getLastRow() >= 3) {
    const o = org.getRange(3, 1, org.getLastRow() - 2, 3).getValues();
    for (let i = 0; i < o.length; i++) {
      if (String(o[i][0] || '').trim().toLowerCase() === cheie && String(o[i][2] || '').trim()) return String(o[i][2]).trim();
    }
  }
  return '';
}

/* ---------- mail „evenimentul tău e pe site” ---------- */
/* La fiecare 15 minute: evenimentele venite prin formular (Adăugat de = Organizator) care au Status = Publicat
   și n-au primit încă mailul îl primesc o singură dată, la adresa din Contact organizator.
   Ca linkul să meargă sigur, mailul pleacă abia la a doua verificare după publicare (site-ul se actualizează în câteva minute).
   Mailul trimis se notează în „Alte detalii / note interne”. Pornire: rulează o dată porneste_mail_evenimente din editor. */
const EV_COL_LINK = 1, EV_COL_CONTACT = 18, EV_COL_NOTE = 19, EV_COL_DE_CINE = 20;
const MARCA_EV_LIVE = 'mail „eveniment pe site” trimis';

function porneste_mail_evenimente() {
  const exista = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'mail_evenimente_publicate'; });
  if (!exista) ScriptApp.newTrigger('mail_evenimente_publicate').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Gata: mailurile „evenimentul e pe site” pleacă singure.');
}

/** Test: trimite pe adresa ta mailul „eveniment publicat”, cu primul eveniment publicat din calendar. */
function testMailEvenimentPublicat() {
  const fila = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Evenimente');
  const v = fila.getRange(3, 1, Math.max(fila.getLastRow() - 2, 1), EV_COL.nume).getValues();
  const r = v.filter(function (x) { return String(x[0] || '').trim(); }).pop() || ['https://cefacemcucopiii.ro/#exemplu', '', '', '', '', '', '', 'Eveniment de test'];
  const rez = trimite_confirmare('eveniment publicat', { 'E-mail': Session.getActiveUser().getEmail(), 'Nume eveniment': String(r[EV_COL.nume - 1]), _linkEv: String(r[0]) });
  SpreadsheetApp.getActiveSpreadsheet().toast('Mail de test „eveniment publicat”: ' + (rez || 'nu a plecat'));
}

function mail_evenimente_publicate() {
  const fila = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Evenimente');
  if (!fila || fila.getLastRow() < 3) return;
  const n = fila.getLastRow() - 2;
  const v = fila.getRange(3, 1, n, EV_COL.ordine).getValues();
  const props = PropertiesService.getScriptProperties();
  const asteapta = JSON.parse(props.getProperty('EV_LIVE_ASTEAPTA') || '{}');
  const acum = Date.now(), nouAsteapta = {};
  const tz = 'Europe/Bucharest';
  v.forEach(function (r, i) {
    const link = String(r[EV_COL_LINK - 1] || '').trim();
    if (!/^publicat/i.test(String(r[EV_COL.status - 1] || '')) || !link) return;
    if (!/^organizator/i.test(String(r[EV_COL_DE_CINE - 1] || '').trim())) return;
    const note = String(r[EV_COL_NOTE - 1] || '');
    if (note.indexOf(MARCA_EV_LIVE) >= 0) return;
    const ord = r[EV_COL.ordine - 1];
    if (ord instanceof Date && ord.getTime() < acum) return; // a trecut deja
    const m = String(r[EV_COL_CONTACT - 1] || '').match(/[^\s,;<>]+@[^\s,;<>]+\.[a-z]{2,}/i);
    if (!m) return;
    // prima dată când îl vedem publicat doar îl notăm; mailul pleacă la verificarea următoare
    if (!asteapta[link]) { nouAsteapta[link] = acum; return; }
    if (acum - asteapta[link] < 8 * 60 * 1000) { nouAsteapta[link] = asteapta[link]; return; }
    const c = { 'E-mail': m[0], 'Nume eveniment': String(r[EV_COL.nume - 1] || '').trim(), _linkEv: link };
    let rez;
    try { rez = trimite_confirmare('eveniment publicat', c); } catch (err) { rez = 'eroare: ' + err; }
    const azi = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
    const nota = rez === 'Brevo' || rez === 'Gmail' ? azi + ' - ' + MARCA_EV_LIVE + ' la ' + m[0]
      : azi + ' - mail „eveniment pe site” NU a plecat la ' + m[0] + (rez ? ' (' + String(rez).slice(0, 150) + ')' : '') + '. Încercăm din nou peste 15 minute.';
    if (rez === 'Brevo' || rez === 'Gmail') fila.getRange(i + 3, EV_COL_NOTE).setValue((note.trim() ? note.trim() + ' | ' : '') + nota);
    else nouAsteapta[link] = asteapta[link];
    try {
      const rp = rand_propunere(fila.getParent(), r[EV_COL.nume - 1], r[EV_COL.data - 1]);
      if (rp) noteaza_propunere(fila.getParent().getSheetByName(FILE.eveniment.nume), rp, nota.replace(/^\S+ \S+ - /, ''));
    } catch (errP) { console.warn(errP); }
  });
  props.setProperty('EV_LIVE_ASTEAPTA', JSON.stringify(nouAsteapta));
}

/** O singură dată, din editor (9 oct): reorganizează coloanele din Evenimente în ordinea nouă. Nu mai trebuie rulată. */
function reorganizeaza_evenimente() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Evenimente');
  const cap = function () { return sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); }); };
  const col = function (inceput) { const h = cap(); for (let i = 0; i < h.length; i++) if (h[i].indexOf(inceput) === 0) return i + 1; return 0; };
  if (!col('Interior') && !col('Note interne')) { SpreadsheetApp.getActiveSpreadsheet().toast('Coloanele sunt deja reorganizate.'); return; }
  const n = sh.getLastRow() - 2;
  // 1. Note interne se adaugă la Alte detalii (nu se pierde nimic)
  const cA = col('Alte detalii'), cN = col('Note interne');
  if (cA && cN && n > 0) {
    const a = sh.getRange(3, cA, n, 1).getValues(), b = sh.getRange(3, cN, n, 1).getValues();
    sh.getRange(3, cA, n, 1).setValues(a.map(function (x, i) { return [[String(x[0] || '').trim(), String(b[i][0] || '').trim()].filter(String).join(' | ')]; }));
  }
  // 2. coloanele de care nu mai avem nevoie
  ['Instagram organizator', 'Conținut dedicat', 'Note interne', 'Unde am găsit', 'Recurență', 'Limbă', 'Rezervare', 'Interior'].forEach(function (h) {
    const c = col(h); if (c) sh.deleteColumn(c);
  });
  // 3. ordinea nouă
  const ordine = ['ID', 'Status', 'Recomandare', 'Data', 'Zi', 'Ora start', 'Ora sfârșit', 'Nume eveniment', 'Organizator', 'Locație', 'Descriere',
    'Vârstă de la', 'Vârstă până la', 'Categorie', 'Adresă', 'Bilet', 'Preț', 'Link bilete', 'Link eveniment', 'Contact organizator', 'Alte detalii',
    'Adăugat de', 'Data adăugării', 'Ordine'];
  sh.getRange(2, 2).setValue('Status');
  ordine.forEach(function (h, p) {
    const c = col(h);
    if (c && c !== p + 1) sh.moveColumns(sh.getRange(1, c, sh.getMaxRows(), 1), p + 1);
  });
  // 4. numele noi din capul de tabel
  sh.getRange(2, 1, 1, 24).setValues([['ID - link pe site', 'Status', 'Recomandare', 'Data', 'Zi', 'Ora start', 'Ora sfârșit', 'Nume eveniment', 'Organizator',
    'Locație', 'Descriere', 'Vârstă de la', 'Vârstă până la', 'Categorie', 'Adresă', 'Bilet', 'Preț (lei)', 'Link bilete', 'Link sursă (doar pentru noi)',
    'Contact organizator', 'Alte detalii / note interne', 'Adăugat de', 'Data adăugării', 'Ordine (pt sortare)']]);
  // 5. Recomandare: Da / Nu
  if (n > 0) {
    const rec = sh.getRange(3, 3, n, 1).getValues(), dat = sh.getRange(3, 4, n, 1).getValues();
    sh.getRange(3, 3, n, 1).setValues(rec.map(function (x, i) { return [String(x[0] || '').trim() || (String(dat[i][0] || '').trim() ? 'Nu' : '')]; }));
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Gata: coloanele din Evenimente sunt în ordinea nouă.');
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
const STATUS_CREATA = 'Pagină creată (Pe site = Nu)'; // status vechi, nu mai e în listă; îl tratăm la fel ca STATUS_CREEAZA

/** Rulează o singură dată din editor: pornește automatizarea din sheet-ul de onboarding. */
function porneste_onboarding() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const h = t.getHandlerFunction();
    if (h === 'la_editare_onboarding' || h === 'la_editare_comunitate') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('la_editare_onboarding').forSpreadsheet(ONBOARDING_ID).onEdit().create();
}

/* Fluxul din Onboarding (totul se face aici; în Comunitatea CFCC ajunge doar ce e publicat):
   3. Verificat - creează pagina  -> link de previzualizare în Note (pagina NU se copiază în Comunitatea CFCC)
   5. Trimis spre aprobare        -> mail către partener cu linkul de previzualizare
   6. Publicat                    -> pagina se copiază în Comunitatea CFCC cu Pe site = Da + mail „pagina e live” */
function la_editare_onboarding(e) {
  const fila = e.range.getSheet();
  if (fila.getName() !== FILE.pagina.nume) return;
  if (e.range.getColumn() !== 2 || e.range.getRow() < 2 || e.range.getNumRows() > 1) return; // coloana Status
  const val = String(e.value || '');
  const rand = e.range.getRow();
  const nume = fila.getRange(rand, 3).getDisplayValue().trim();
  const id = slug_id(nume);
  const titlu = 'Ce facem cu copiii?';
  coloreaza_rand(fila, rand, val);
  try {
    if (val === STATUS_CREEAZA || val === STATUS_CREATA) {
      if (!nume) throw new Error('lipsește numele');
      const link = link_previzualizare(fila, rand, true);
      e.source.toast('Pagina „' + nume + '” e gata de verificat. Linkul de previzualizare e în coloana Note: ' + link, titlu, 15);
    } else if (/aprobare/i.test(val)) {
      const link = link_previzualizare(fila, rand, true);
      e.source.toast(mail_spre_aprobare(fila, rand, id, link), titlu, 12);
    } else if (/^publicat/i.test(val)) {
      const idPub = publica_pagina(fila, rand);
      e.source.toast('Pagina e în Comunitatea CFCC (cefacemcucopiii.ro/#comunitate/' + idPub + '), cu Pe site = Da. Apare pe site în câteva minute. ' + mail_pagina_live(fila, rand, idPub), titlu, 15);
    }
  } catch (err) {
    e.source.toast('Nu am putut termina pasul: ' + err.message, titlu, 15);
  }
}

/** Culoarea rândului după status: roșu = la noi, în verificare; portocaliu = trimis spre aprobare; verde = publicat; gri = respins. */
function coloreaza_rand(fila, rand, status) {
  const c = /^publicat/i.test(status) ? ['#d9ead3', '#274e13']
    : /aprobare/i.test(status) ? ['#fce5cd', '#7f4f00']
    : /respins/i.test(status) ? ['#efefef', '#666666']
    : ['#f4c7c3', '#9c0006'];
  fila.getRange(rand, 1, 1, FILE.pagina.coloane.length).setBackground(c[0]).setFontColor(c[1]);
}

/** Linkul de previzualizare (cu un cod greu de ghicit), notat în coloana Note. Dacă există deja, îl refolosește. */
function link_previzualizare(fila, rand, creeaza) {
  const cols = FILE.pagina.coloane.map(function (c) { return c[0]; });
  const colNote = cols.indexOf('Note') + 1;
  const nota = fila.getRange(rand, colNote).getDisplayValue();
  const m = nota.match(/#previzualizare\/([a-f0-9]{16,})/);
  if (m) return SITE + '/#previzualizare/' + m[1];
  if (!creeaza) return '';
  const cod = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  const link = SITE + '/#previzualizare/' + cod;
  // facem pozele vizibile pentru cine are linkul, ca să apară în previzualizare
  ['Logo', 'Poza principală', 'Alte poze'].forEach(function (k) {
    try { poze_publice(fila.getRange(rand, cols.indexOf(k) + 1).getDisplayValue()); } catch (err) { console.warn('Poze: ' + err); }
  });
  const azi = Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy HH:mm');
  fila.getRange(rand, colNote).setValue((nota ? nota.replace(/\s+$/, '') + '\n\n' : '') + azi + ' - previzualizare (nu e publică): ' + link);
  return link;
}

/** Datele publice ale unei pagini în previzualizare (pentru site). Doar câmpurile care apar pe pagină, niciodată datele firmei. */
function date_previzualizare(cod) {
  if (!/^[a-f0-9]{16,}$/.test(cod || '')) return { ok: false };
  const fila = SpreadsheetApp.openById(ONBOARDING_ID).getSheetByName(FILE.pagina.nume);
  if (!fila || fila.getLastRow() < 2) return { ok: false };
  const cols = FILE.pagina.coloane.map(function (c) { return c[0]; });
  const rows = fila.getRange(2, 1, fila.getLastRow() - 1, cols.length).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    const p = {};
    cols.forEach(function (k, j) { p[k] = String(rows[i][j] || '').trim(); });
    if (p['Note'].indexOf('#previzualizare/' + cod) < 0) continue;
    if (/respins/i.test(p['Status'])) return { ok: false };
    const poze = function (k) { return linkuri_poze(p[k]); };
    const par = SpreadsheetApp.openById(COMUNITATE_ID).getSheetByName('Parteneri');
    const existent = pagina_existenta(par, p);
    const r = rand_comunitate(p, existent, true);
    const vechi_act = [], vechi_foto = [];
    if (existent && (!p['Activități'].replace(/[\s|]/g, '') || !p['Alte poze'])) {
      const com = SpreadsheetApp.openById(COMUNITATE_ID);
      if (!p['Activități'].replace(/[\s|]/g, '')) { const fa = com.getSheetByName('Activități'); if (fa.getLastRow() > 1) fa.getRange(2, 1, fa.getLastRow() - 1, 3).getDisplayValues().forEach(function (x) { if (x[0] === r[0] && x[1]) vechi_act.push([x[1], x[2]]); }); }
      if (!p['Alte poze']) { const fg = com.getSheetByName('Galerie'); if (fg.getLastRow() > 1) fg.getRange(2, 1, fg.getLastRow() - 1, 2).getDisplayValues().forEach(function (x) { if (x[0] === r[0] && x[1]) vechi_foto.push(x[1]); }); }
    }
    const act = p['Activități'].split(/\n+/).filter(function (l) { return l.replace(/\|/g, '').trim(); }).map(function (l) { const x = l.split('|'); return [x[0].trim(), (x.slice(1).join('|') || '').trim()]; });
    return { ok: true, pagina: {
      id: r[0], n: r[1], tip: r[2], pro: /^partener/i.test(r[3]), full: /^partener|complet/i.test(norm_nume(r[3])), inlocuieste: existent ? r[0] : '', desc: r[5], age: r[6], ad: r[7],
      tel: String(r[8]).replace(/^'/, ''), mail: r[9], ig: r[10], fb: r[11], web: r[12], logo: r[13], cover: r[14],
      foto: p['Alte poze'] ? poze('Alte poze') : vechi_foto, nume: r[15].split('\n'), act: act.length ? act : vechi_act
    } };
  }
  return { ok: false };
}

function doGet(e) {
  const cod = e && e.parameter && e.parameter.previzualizare;
  if (cod) return raspuns(date_previzualizare(String(cod)));
  return raspuns({ ok: true });
}

/* ---- legătura dintre un rând de onboarding și Comunitatea CFCC ---- */

function norm_nume(x) { return String(x || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

/** Câmpurile publice din rândul de onboarding. */
function campuri_onboarding(fila, rand) {
  const v = fila.getRange(rand, 1, 1, FILE.pagina.coloane.length).getDisplayValues()[0];
  const p = {};
  FILE.pagina.coloane.forEach(function (c, i) { p[c[0]] = String(v[i] || '').trim(); });
  return p;
}

/** Caută în Comunitatea CFCC o pagină care există deja pentru același loc (după ID, nume sau „Nume în calendar”).
    Întoarce { rand, valori } sau null. Așa, un partener care avea deja pagină (ex. o instituție) o primește pe cea nouă în loc, cu același link. */
function pagina_existenta(par, p) {
  if (par.getLastRow() < 2) return null;
  const rows = par.getRange(2, 1, par.getLastRow() - 1, 16).getDisplayValues();
  const id = slug_id(p['Nume']);
  const v = norm_nume(p['Nume']);
  const noi = [v].concat(String(p['Nume în calendar'] || '').split(/\n+/).map(norm_nume)).filter(String);
  let best = null, scor = 0;
  rows.forEach(function (r, i) {
    if (r[0] === id) { best = { rand: i + 2, valori: r }; scor = 1e9; return; }
    const chei = [r[1]].concat(String(r[15] || '').split(/\n+/)).map(norm_nume).filter(String);
    chei.forEach(function (k) {
      noi.forEach(function (n) {
        const ok = n === k || (k.length >= 8 && (' ' + n + ' ').indexOf(' ' + k + ' ') >= 0) || (n.length >= 8 && (' ' + k + ' ').indexOf(' ' + n + ' ') >= 0);
        if (ok && Math.min(k.length, n.length) > scor) { scor = Math.min(k.length, n.length); best = { rand: i + 2, valori: r }; }
      });
    });
  });
  return best;
}

/** Rândul pentru „Parteneri” din Comunitatea CFCC, din datele de onboarding. Ce lipsește din formular rămâne ca în pagina veche. */
function rand_comunitate(p, existent, doar_citire) {
  const vechi = existent ? existent.valori : [];
  const ia = function (nou, i) { return nou ? nou : (vechi[i] || ''); };
  const poze = doar_citire ? linkuri_poze : poze_publice; // la previzualizare nu schimbăm permisiunile din Drive
  const logo = poze(p['Logo'])[0] || '';
  const cover = poze(p['Poza principală'])[0] || '';
  const chei = [];
  [p['Nume']].concat(String(p['Nume în calendar'] || '').split(/\n+/), String(vechi[15] || '').split(/\n+/)).forEach(function (x) {
    x = String(x).trim(); if (x && chei.map(norm_nume).indexOf(norm_nume(x)) < 0) chei.push(x);
  });
  return [existent ? vechi[0] : slug_id(p['Nume']), p['Nume'], ia(p['Cum se descriu'], 2), existent ? (vechi[3] || 'Partener') : 'Partener', 'Da',
    ia(p['Descriere'], 5), ia(p['Vârste'], 6), ia(p['Adresă'], 7), p['Telefon public'] ? text(p['Telefon public']) : (vechi[8] ? text(vechi[8]) : ''),
    ia(p['E-mail public'], 9), ia(link_social(p['Instagram'], 'instagram'), 10), ia(link_social(p['Facebook'], 'facebook'), 11),
    ia(p['Website'], 12), ia(logo, 13), ia(cover, 14), chei.join('\n'),
    SITE + '/#comunitate/' + (existent ? vechi[0] : slug_id(p['Nume']))];
}

/** La „Publicat”: pune pagina în Comunitatea CFCC cu Pe site = Da. Dacă locul avea deja pagină, o înlocuiește (același ID, același link). Întoarce ID-ul. */
function publica_pagina(fila, rand) {
  const p = campuri_onboarding(fila, rand);
  if (!p['Nume']) throw new Error('lipsește numele');
  const com = SpreadsheetApp.openById(COMUNITATE_ID);
  const par = com.getSheetByName('Parteneri');
  const existent = pagina_existenta(par, p);
  const r = rand_comunitate(p, existent);
  const id = r[0];
  if (existent) par.getRange(existent.rand, 1, 1, r.length).setValues([r]);
  else par.appendRow(r);
  const act = String(p['Activități'] || '').split(/\n+/).filter(function (l) { return l.replace(/\|/g, '').trim(); }).map(function (l) {
    const x = l.split('|'); return [id, x[0].trim(), (x.slice(1).join('|') || '').trim()];
  });
  const alte = poze_publice(p['Alte poze']);
  // activitățile și pozele vechi se înlocuiesc doar dacă au venit altele noi
  if (act.length) { const fa = com.getSheetByName('Activități'); sterge_ramase(fa, id); fa.getRange(fa.getLastRow() + 1, 1, act.length, 3).setValues(act); }
  if (alte.length) { const fg = com.getSheetByName('Galerie'); sterge_ramase(fg, id); fg.getRange(fg.getLastRow() + 1, 1, alte.length, 3).setValues(alte.map(function (u) { return [id, u, '']; })); }
  return id;
}

/** Trimite partenerului un mail o singură dată și îl notează în coloana Note. */
function mail_o_data(fila, rand, id, tip, eticheta, link) {
  const nr = FILE.pagina.coloane.length;
  const v = fila.getRange(rand, 1, 1, nr).getDisplayValues()[0];
  const p = {};
  FILE.pagina.coloane.forEach(function (c, i) { p[c[0]] = String(v[i] || '').trim(); });
  const colNote = FILE.pagina.coloane.map(function (c) { return c[0]; }).indexOf('Note') + 1;
  if (p['Note'].indexOf('mail „' + eticheta + '” trimis') >= 0) return 'Mailul „' + eticheta + '” fusese deja trimis, nu l-am trimis din nou.';
  if (!/^\S+@\S+\.\S+$/.test(p['E-mail'])) return 'Nu am trimis mailul „' + eticheta + '”: lipsește e-mailul persoanei de contact.';
  trimite_confirmare(tip, { 'E-mail': p['E-mail'], 'Persoană de contact': p['Persoană de contact'], 'Nume': p['Nume'], _id: id, _link: link });
  const azi = Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy HH:mm');
  fila.getRange(rand, colNote).setValue((p['Note'] ? p['Note'].replace(/\s+$/, '') + '\n\n' : '') + azi + ' - mail „' + eticheta + '” trimis la ' + p['E-mail']);
  return 'Am trimis mailul „' + eticheta + '” la ' + p['E-mail'] + '.';
}

/** Mailul „pagina e live”, o singură dată; pornește și verificarea zilnică pentru mailul cu recomandări. */
function mail_pagina_live(fila, rand, id) {
  const r = mail_o_data(fila, rand, id, 'pagina live', 'pagina e live');
  try { asigura_verificare_zilnica(); } catch (err) { console.warn('Trigger zilnic: ' + err); }
  return r;
}

/** Mailul cu linkul de previzualizare, trimis spre aprobare, o singură dată. */
function mail_spre_aprobare(fila, rand, id, link) {
  return mail_o_data(fila, rand, id, 'pagina spre aprobare', 'spre aprobare', link);
}

/** Pornește (o singură dată) verificarea zilnică de la 10:00 pentru mailurile programate. */
function asigura_verificare_zilnica() {
  const exista = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'mailuri_programate'; });
  if (!exista) ScriptApp.newTrigger('mailuri_programate').timeBased().everyDays(1).atHour(10).inTimezone('Europe/Bucharest').create();
}

/** Rulează zilnic: partenerilor publicați de cel puțin ~20 de ore le trimite mailul cu recomandări (o singură dată). */
function mailuri_programate() {
  const fila = SpreadsheetApp.openById(ONBOARDING_ID).getSheetByName(FILE.pagina.nume);
  if (!fila || fila.getLastRow() < 2) return;
  const cols = FILE.pagina.coloane.map(function (c) { return c[0]; });
  const colNote = cols.indexOf('Note') + 1;
  const rows = fila.getRange(2, 1, fila.getLastRow() - 1, cols.length).getDisplayValues();
  rows.forEach(function (r, i) {
    const p = {};
    cols.forEach(function (k, j) { p[k] = String(r[j] || '').trim(); });
    if (!/^publicat/i.test(p['Status'])) return;
    const m = p['Note'].match(/(\d{2})\.(\d{2})\.(\d{4}) (\d{2}):(\d{2}) - mail „pagina e live” trimis/);
    if (!m) return;
    const live = new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]);
    const ore = (Date.now() - live.getTime()) / 3600000;
    const azi = Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy HH:mm');
    // pasul 2 (după ~1 zi): mail cu recomandări către partener
    if (ore >= 20 && !/mail „recomandări” trimis/i.test(p['Note']) && /^\S+@\S+\.\S+$/.test(p['E-mail'])) {
      trimite_confirmare('pagina live - recomandări', { 'E-mail': p['E-mail'], 'Persoană de contact': p['Persoană de contact'], 'Nume': p['Nume'], _id: slug_id(p['Nume']) });
      p['Note'] += (p['Note'] ? '\n\n' : '') + azi + ' - mail „recomandări” trimis la ' + p['E-mail'];
      fila.getRange(i + 2, colNote).setValue(p['Note']);
    }
  });
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

/** Ca poze_publice, dar doar transformă linkurile din Drive, fără să schimbe permisiunile. */
function linkuri_poze(celula) {
  return String(celula || '').split(/\s+/).filter(String).map(function (u) {
    const m = u.match(/\/d\/([\w-]{20,})/) || u.match(/[?&]id=([\w-]{20,})/);
    return m ? 'https://lh3.googleusercontent.com/d/' + m[1] : u;
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
