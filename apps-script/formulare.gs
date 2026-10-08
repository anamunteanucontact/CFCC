/**
 * Ce facem cu copiii? - primește formularele de pe site, le scrie în sheet
 * și trimite un e-mail de confirmare celui care a completat.
 * Fiecare tip de formular are fila lui: Propuneri evenimente, Colaborări, Abonați, Pagini parteneri - de verificat.
 * Paginile trimise de parteneri intră cu roșu, ca să le verifice un om înainte să ajungă în sheet-ul Comunitatea CFCC.
 * Pozele lor se salvează în Drive, în folderul „CFCC - poze parteneri”.
 * Filele se creează singure la primul mesaj, cu capul de tabel de mai jos.
 */

const NUME_EXPEDITOR = 'Ce facem cu copiii?';
const SITE = 'https://cefacemcucopiii.ro';
const INSTAGRAM = 'https://www.instagram.com/cefacemcucopiii/';

const FILE = {
  eveniment: {
    nume: 'Propuneri evenimente',
    coloane: ['Primit la', 'Nume eveniment', 'Organizator', 'Data', 'Ora', 'Locație', 'Adresă',
      'Vârstă minimă', 'Vârstă maximă', 'Acces', 'Link bilete / înscriere', 'Descriere', 'E-mail', 'Telefon', 'Status']
  },
  colaborare: {
    nume: 'Colaborări',
    coloane: ['Primit la', 'Afacere / brand', 'Persoană de contact', 'E-mail', 'Telefon', 'Instagram',
      'Facebook', 'Website', 'Despre ei', 'Ce au nevoie', 'Status']
  },
  abonare: {
    nume: 'Abonați',
    coloane: ['Primit la', 'E-mail', 'Vrea oferte de la parteneri']
  },
  pagina: {
    nume: 'Pagini parteneri - de verificat',
    coloane: ['Primit la', 'Status', 'Nume', 'Ce sunt', 'Descriere', 'Vârste', 'Adresă', 'Telefon public', 'E-mail public',
      'Instagram', 'Facebook', 'Website', 'Activități', 'Nume în calendar', 'Logo', 'Poza principală', 'Alte poze',
      'Folder poze', 'Persoană de contact', 'Telefon contact', 'E-mail', 'Observații', 'Acord publicare'],
    status: 'De verificat',
    rosu: true,
    lungime: 5000
  }
};

const FOLDER_POZE = 'CFCC - poze parteneri';

function doPost(e) {
  try {
    const date = JSON.parse(e.postData.contents);
    const tip = FILE[date.tip];
    if (!tip) return raspuns({ ok: false, eroare: 'tip necunoscut' });

    const campuri = date.campuri || {};
    if (date.tip === 'pagina') salveaza_poze(campuri);

    const fila = ia_fila(tip);
    const rand = tip.coloane.map(function (col) {
      if (col === 'Primit la') return new Date();
      if (col === 'Status') return tip.status || 'Nou';
      const v = campuri[col];
      return v === undefined ? '' : String(v).slice(0, tip.lungime || 2000);
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
      '<p>Mulțumim că ne-ai trimis <strong>' + esc(c['Nume eveniment']) + '</strong>. Îl verificăm și îl adăugăm în calendarul de pe <a href="' + SITE + '">cefacemcucopiii.ro</a>. Programul nou apare în fiecare săptămână.</p>' +
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
      '<p>Mulțumim! Am primit informațiile pentru pagina <strong>' + esc(c['Nume']) + '</strong> din Comunitatea CFCC. Le citim, aranjăm textele dacă e nevoie și vă trimitem pagina spre aprobare înainte s-o publicăm.</p>' +
      '<p>Dacă vreți să schimbați ceva între timp, răspundeți la acest e-mail.</p>' +
      '<p>Cu drag,<br>Ana, de la Ce facem cu copiii?</p>';
  } else {
    return;
  }

  MailApp.sendEmail({ to: catre, subject: subiect, htmlBody: continut, name: NUME_EXPEDITOR });
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

function esc(s) {
  return String(s || '').replace(/[&<>"]/g, function (ch) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch];
  });
}

function ia_fila(tip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fila = ss.getSheetByName(tip.nume);
  if (!fila) {
    fila = ss.insertSheet(tip.nume);
    fila.appendRow(tip.coloane);
    fila.setFrozenRows(1);
    fila.getRange(1, 1, 1, tip.coloane.length).setFontWeight('bold');
    if (tip.rosu) fila.getRange(1, 1, 1, tip.coloane.length).setBackground('#cc0000').setFontColor('#ffffff');
  }
  return fila;
}

function raspuns(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Rulează o dată din editor ca să aprobi trimiterea de e-mailuri. Îți trimite un mail de probă ție. */
function testMail() {
  trimite_confirmare('abonare', { 'E-mail': Session.getActiveUser().getEmail(), 'Vrea oferte de la parteneri': 'Da' });
}
