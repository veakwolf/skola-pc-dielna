// PC Dielňa – príjem výsledkov do Google Tabuľky
// Vlož do Rozšírenia → Apps Script v tabuľke a nasaď ako webovú aplikáciu.

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const odov = sheet_(ss, "Odovzdania", [
    "Odoslané", "Meno", "Priezvisko", "Trieda", "Začiatok", "Koniec", "Minút",
    "Zostavy (z 3)", "Montáž hotová", "Montáž chyby", "BIOS správne", "BIOS spolu",
    "Servis správne", "Servis spolu", "Chybných krokov spolu"
  ]);
  odov.appendRow([
    new Date(), d.meno, d.priezvisko, d.trieda, new Date(d.zaciatok), new Date(d.koniec), d.minut,
    d.zostavy, d.montaz_hotova ? "áno" : "nie", d.montaz_chyby, d.bios, d.bios_spolu,
    d.servis, d.servis_spolu, d.chyby_spolu
  ]);

  const kroky = sheet_(ss, "Kroky", [
    "Meno", "Priezvisko", "Trieda", "Čas", "Sekunda od začiatku", "Časť", "Akcia", "Detail", "Správne"
  ]);
  const rows = (d.kroky || []).map(k => [
    d.meno, d.priezvisko, d.trieda, new Date(k.t), k.sec, k.cast, k.akcia, k.detail, k.ok
  ]);
  if (rows.length) kroky.getRange(kroky.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  return ContentService.createTextOutput("ok");
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight("bold");
  }
  return sh;
}
