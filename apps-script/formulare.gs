/**
 * Ce facem cu copiii? - primește formularele de pe site și le scrie în sheet.
 * Fiecare tip de formular are fila lui: Propuneri evenimente, Colaborări, Abonați.
 * Filele se creează singure la primul mesaj, cu capul de tabel de mai jos.
 */

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
  }
};

function doPost(e) {
  try {
    const date = JSON.parse(e.postData.contents);
    const tip = FILE[date.tip];
    if (!tip) return raspuns({ ok: false, eroare: 'tip necunoscut' });

    const fila = ia_fila(tip);
    const rand = tip.coloane.map(function (col) {
      if (col === 'Primit la') return new Date();
      if (col === 'Status') return 'Nou';
      const v = date.campuri[col];
      return v === undefined ? '' : String(v).slice(0, 2000);
    });
    fila.appendRow(rand);
    return raspuns({ ok: true });
  } catch (err) {
    return raspuns({ ok: false, eroare: String(err) });
  }
}

function ia_fila(tip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fila = ss.getSheetByName(tip.nume);
  if (!fila) {
    fila = ss.insertSheet(tip.nume);
    fila.appendRow(tip.coloane);
    fila.setFrozenRows(1);
    fila.getRange(1, 1, 1, tip.coloane.length).setFontWeight('bold');
  }
  return fila;
}

function raspuns(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
