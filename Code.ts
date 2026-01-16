// TODO: Terminbeginn in terminNTD

interface MapS2I {
  [others: string]: number;
}
interface MapS2S {
  [others: string]: string;
}
interface HeaderMap {
  [others: string]: MapS2I;
}
let inited = false;
let headers: HeaderMap = {};
let termineSheet: GoogleAppsScript.Spreadsheet.Sheet;
let anmeldungenSheet: GoogleAppsScript.Spreadsheet.Sheet;
let stornierungenSheet: GoogleAppsScript.Spreadsheet.Sheet;

// Indices are 1-based!!
// Anmeldungen:
let mailIndex: number; // E-Mail-Adresse
let vornameIndex: number; // Vorname
let nachnameIndex: number; // Nachname
let adresseIndex: number; // Adresse
let telefonIndex: number; // Telefonnummer
let anmeldebestIndex: number; // Anmeldebestätigung (gesendet)
let anmeldeTerminIndex: number; // Termin

// Termine:
let terminIndex: number; // Termin
let terminPlätzeIndex: number; // Plätze
let restPlätzeIndex: number; // Rest

// Stornierungen:
let stornoMailIndex: number; // Email
let stornoTerminIndex: number; // Termin

// map Anmeldungen headers to print headers
let printCols = new Map([
  ["Vorname", "Vorname"],
  ["Nachname", "Nachname"],
  ["E-Mail-Adresse", "Email"],
  ["Telefonnummer", "Telefon"],
]);

const terminFrage = "Termin";

interface SSEvent {
  namedValues: { [others: string]: string[] };
  range: GoogleAppsScript.Spreadsheet.Range;
  [others: string]: any;
}

function isEmpty(str: string | undefined | null) {
  if (typeof str == "number") return false;
  return !str || 0 === str.length; // I think !str is sufficient...
}

function init() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  for (let sheet of sheets) {
    let sheetName = sheet.getName();
    let sheetHeaders: MapS2I = {};
    // Logger.log("sheetName %s", sheetName);
    headers[sheetName] = sheetHeaders;
    let numCols = sheet.getLastColumn();
    // Logger.log("numCols %s", numCols);
    let row1Vals = sheet.getRange(1, 1, 1, numCols).getValues();
    // Logger.log("sheetName %s row1 %s", sheetName, row1Vals);
    for (let i = 0; i < numCols; i++) {
      let v: string = row1Vals[0][i];
      if (isEmpty(v)) continue;
      sheetHeaders[v] = i + 1;
    }
    // Logger.log("sheet %s %s", sheetName, sheetHeaders);

    if (sheet.getName() == "Termine") {
      termineSheet = sheet;
      terminIndex = sheetHeaders["Termin"];
      terminPlätzeIndex = sheetHeaders["Plätze"];
      restPlätzeIndex = sheetHeaders["Rest"];
    }
    if (sheet.getName() == "Anmeldungen") {
      anmeldungenSheet = sheet;
      mailIndex = sheetHeaders["E-Mail-Adresse"];
      vornameIndex = sheetHeaders["Vorname"];
      nachnameIndex = sheetHeaders["Nachname"];
      adresseIndex = sheetHeaders["Adresse"];
      telefonIndex = sheetHeaders["Telefonnummer"];
      anmeldeTerminIndex = sheetHeaders[terminFrage];
      anmeldebestIndex = sheetHeaders["Anmeldebestätigung"];
      if (anmeldebestIndex == null) {
        anmeldebestIndex = addColumn(sheet, sheetHeaders, "Anmeldebestätigung");
      }
    }
    if (sheet.getName() == "Stornierungen") {
      stornierungenSheet = sheet;
      stornoMailIndex = sheetHeaders["Email"];
      stornoTerminIndex = sheetHeaders["Termin"];
    }
  }
  inited = true;
}

// add a cell in row 1 with a new column title, return its index
function addColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetHeaders: MapS2I,
  title: string,
): number {
  let max = 0;
  for (let sh in sheetHeaders) {
    if (sheetHeaders[sh] > max) max = sheetHeaders[sh];
  }
  if (max >= sheet.getMaxColumns()) {
    sheet.insertColumnAfter(max);
  }
  max += 1;
  sheet.getRange(1, max).setValue(title);
  sheetHeaders[title] = max;
  return max;
}

function anredeText(vorname: string, nachname: string) {
  return "Liebe(r) " + vorname + " " + nachname;
}

function heuteString() {
  return Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    "YYYY-MM-dd HH:mm:ss",
  );
}

function anmeldebestätigungFromMenu() {
  Logger.log("anmeldebestätigung");
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Anmeldungen") {
    SpreadsheetApp.getUi().alert("Bitte eine Zeile im Sheet 'Anmeldungen' selektieren");
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert("Bitte zuerst Teilnehmerzeile selektieren");
    return;
  }
  let row = curCell.getRow();
  if (row < 2 || row > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren",
    );
    return;
  }
  let rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  let rowNote = sheet.getRange(row, 1).getNote();
  if (!isEmpty(rowNote)) {
    SpreadsheetApp.getUi().alert("Die ausgewählte Zeile hat eine Notiz und ist deshalb ungültig");
    return;
  }
  if (!isEmpty(rowValues[anmeldebestIndex - 1])) {
    SpreadsheetApp.getUi().alert("Der Termin wurde schon bestätigt");
    return;
  }

  sendeBestätigung(sheet, row, rowValues, true);
}

function sendeBestätigung(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  rowValues: any[],
  erfolg: boolean,
) {
  // setting up mail
  let emailTo: string = rowValues[mailIndex - 1];
  let subject: string = erfolg ? "Bestätigung deiner Anmeldung" : "Leider erfolglose Anmeldung";
  let vorname: string = rowValues[vornameIndex - 1];
  let nachname: string = rowValues[nachnameIndex - 1];
  let addresse: string = rowValues[adresseIndex - 1];

  let anrede: string = anredeText(vorname, nachname);
  let termin: string = rowValues[anmeldeTerminIndex - 1];

  let template: GoogleAppsScript.HTML.HtmlTemplate =
    HtmlService.createTemplateFromFile("emailBestätigung.html");
  template.anrede = anrede;
  template.termin = termin;
  template.erfolg = erfolg;
  template.adresse = addresse;

  let htmlText: string = template.evaluate().getContent();
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "AG Codierung des ADFC München e.V.",
    replyTo: "anmeldungen-codierung@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
  // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());
}

function sendeStornierung(sheet: GoogleAppsScript.Spreadsheet.Sheet, rowValues: any[]) {
  // setting up mail
  let emailTo: string = rowValues[mailIndex - 1];
  let subject: string = "Bestätigung deiner Stornierung";
  let vorname: string = rowValues[vornameIndex - 1];
  let nachname: string = rowValues[nachnameIndex - 1];

  let anrede: string = anredeText(vorname, nachname);
  let termin: string = rowValues[anmeldeTerminIndex - 1];

  let template: GoogleAppsScript.HTML.HtmlTemplate =
    HtmlService.createTemplateFromFile("emailStornierung.html");
  template.anrede = anrede;
  template.termin = termin;

  let htmlText: string = template.evaluate().getContent();
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "AG Codierung des ADFC München e.V.",
    replyTo: "anmeldungen-codierung@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("ADFC-CODIER")
    .addItem("Anmeldebestätigung senden", "anmeldebestätigungFromMenu")
    .addItem("Update", "update")
    .addItem("Terminteilnehmer drucken", "printTerminMembers")
    .addItem("Anmeldung prüfen", "checkAnmeldungManually")
    .addToUi();
}

function dispatch(e: SSEvent) {
  let docLock = LockService.getScriptLock();
  let locked = docLock.tryLock(30000);
  if (!locked) {
    Logger.log("Could not obtain document lock");
  }
  if (!inited) init();
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  Logger.log("dispatch sheet %s %s", sheet.getName(), range.getA1Notation());
  if (sheet.getName() == "Anmeldungen") checkAnmeldung(e);
  if (sheet.getName() == "Stornierungen") checkStornierung(e);
  if (locked) docLock.releaseLock();
}

function checkAnmeldung(e: SSEvent) {
  Logger.log("checkAnmeldung %s", e.namedValues);
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  let row = range.getRow();
  let cellA = range.getCell(1, 1);
  Logger.log("sheet %s row %s cellA %s", sheet, row, cellA.getA1Notation());

  let rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  let termin1 = e.namedValues[terminFrage][0];
  let termin2 = rowValues[anmeldeTerminIndex - 1];
  Logger.log("2checkAnmeldung %s %s", termin1, termin2);

  let termineVals: Array<Array<string>> = termineSheet.getSheetValues(
    2,
    1,
    termineSheet.getLastRow(),
    termineSheet.getLastColumn(),
  );

  let anmeldungenRows = anmeldungenSheet.getLastRow() - 1; // first row = headers
  let anmeldungenCols = anmeldungenSheet.getLastColumn();
  let anmeldungenVals = anmeldungenSheet
    .getRange(2, 1, anmeldungenRows, anmeldungenCols)
    .getValues();
  let anmeldungenNotes = anmeldungenSheet.getRange(2, 1, anmeldungenRows, 1).getNotes();

  for (let b = 0; b < anmeldungenRows; b++) {
    if (b + 2 == row) continue;
    if (!isEmpty(anmeldungenNotes[b][0])) continue;
    if (
      anmeldungenVals[b][anmeldeTerminIndex - 1] === termin1 &&
      anmeldungenVals[b][vornameIndex - 1].trim().toLowerCase() ===
        rowValues[vornameIndex - 1].trim().toLowerCase() &&
      anmeldungenVals[b][nachnameIndex - 1].trim().toLowerCase() ===
        rowValues[nachnameIndex - 1].trim().toLowerCase()
    ) {
      anmeldungenSheet.getRange(row, 1).setNote("Doppelt");
      sendeBestätigung(sheet, row, rowValues, true);
      return;
    }
  }

  let restChanged = false;
  let terminFound = false;
  let erfolg = true;
  for (let j = 0; j < termineVals.length; j++) {
    let termineRow = termineVals[j];
    if (!termineRow[0]) continue;
    let ttermin = termineRow[terminIndex - 1];
    if (termin1 === ttermin) {
      terminFound = true;
      let rest = termineSheet.getRange(2 + j, restPlätzeIndex).getValue();
      if (rest <= 0) {
        Logger.log("Termin '" + termin1 + "' überbucht!");
        erfolg = false;
        sheet.getRange(row, 1).setNote("Überbucht");
      } else {
        termineSheet.getRange(2 + j, restPlätzeIndex).setValue(rest - 1);
        restChanged = true;
      }
      break;
    }
  }
  if (!terminFound) {
    Logger.log("Termin '" + termin1 + "' nicht im Termine-Sheet!?");
    return;
  }
  if (restChanged) {
    updateForm();
  }
  sendeBestätigung(sheet, row, rowValues, erfolg);
}

function checkStornierung(e: SSEvent) {
  Logger.log("1checkStornierung %s", e.namedValues);
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  let row = range.getRow();
  let cellA = range.getCell(1, 1);
  Logger.log("sheet %s row %s cellA %s", sheet, row, cellA.getA1Notation());

  let rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  let termin1 = e.namedValues[terminFrage][0];
  let termin2 = rowValues[stornoTerminIndex - 1];
  Logger.log("2checkStornierung %s %s", termin1, termin2);

  let anmeldungenRows = anmeldungenSheet.getLastRow() - 1; // first row = headers
  let anmeldungenCols = anmeldungenSheet.getLastColumn();
  let anmeldungenVals = anmeldungenSheet
    .getRange(2, 1, anmeldungenRows, anmeldungenCols)
    .getValues();
  let anmeldungenNotes = anmeldungenSheet.getRange(2, 1, anmeldungenRows, 1).getNotes();

  let storniert = false;
  for (let b = 0; b < anmeldungenRows; b++) {
    if (!isEmpty(anmeldungenNotes[b][0])) continue;
    if (
      anmeldungenVals[b][anmeldeTerminIndex - 1] === termin1 &&
      anmeldungenVals[b][mailIndex - 1].trim().toLowerCase() ===
        rowValues[stornoMailIndex - 1].trim().toLowerCase()
    ) {
      storniert = true;
      anmeldungenSheet.getRange(row, 1).setNote("Storniert");
      sendeStornierung(sheet, anmeldungenVals[b]);
      break;
    }
  }
  if (storniert) {
    updateTerminReste();
    updateForm();
  }
}

function update() {
  let docLock = LockService.getScriptLock();
  let locked = docLock.tryLock(30000);
  if (!locked) {
    SpreadsheetApp.getUi().alert("Konnte Dokument nicht locken");
    return;
  }
  if (!inited) init();
  updateTerminReste();
  updateForm();
  docLock.releaseLock();
}

function date2Str(ddate: Date): string {
  let sdate: string = Utilities.formatDate(
    ddate,
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    "dd.MM.YYYY",
  );
  return sdate;
}

function updateTerminReste() {
  let termineRows = termineSheet.getLastRow() - 1; // first row = headers
  let termineCols = termineSheet.getLastColumn();
  let termineVals = termineSheet.getRange(2, 1, termineRows, termineCols).getValues();
  let termineNotes = termineSheet.getRange(2, 1, termineRows, 1).getNotes();

  let anmeldungenRows = anmeldungenSheet.getLastRow() - 1; // first row = headers
  let anmeldungenCols = anmeldungenSheet.getLastColumn();
  let anmeldungenVals: any[][];
  let anmeldungenNotes: string[][];
  // getRange with 0 rows throws an exception instead of returning an empty array
  if (anmeldungenRows == 0) {
    anmeldungenVals = [];
    anmeldungenNotes = [];
  } else {
    anmeldungenVals = anmeldungenSheet.getRange(2, 1, anmeldungenRows, anmeldungenCols).getValues();
    anmeldungenNotes = anmeldungenSheet.getRange(2, 1, anmeldungenRows, 1).getNotes();
  }

  let terminplätze: MapS2I = {};
  for (let b = 0; b < anmeldungenRows; b++) {
    if (!isEmpty(anmeldungenNotes[b][0])) continue;
    let termin = anmeldungenVals[b][anmeldeTerminIndex - 1];
    let anzahl: number = terminplätze[termin];
    if (anzahl == null) {
      terminplätze[termin] = 1;
    } else {
      terminplätze[termin] = anzahl + 1;
    }
  }

  for (let r = 0; r < termineRows; r++) {
    if (!isEmpty(termineNotes[r][0])) continue;
    let termin: string = termineVals[r][terminIndex - 1];
    let terminPlätze: number = termineVals[r][terminPlätzeIndex - 1];
    let restPlätze: number = termineVals[r][restPlätzeIndex - 1];

    let terminGebucht: number = terminplätze[termin];
    if (terminGebucht == null) terminGebucht = 0;
    let terminRest: number = terminPlätze - terminGebucht;
    if (terminRest < 0) {
      SpreadsheetApp.getUi().alert("Der Termin '" + termin + "' ist überbucht!");
      terminRest = 0;
    }
    if (terminRest !== restPlätze) {
      termineSheet.getRange(2 + r, restPlätzeIndex).setValue(terminRest);
      SpreadsheetApp.getUi().alert(
        "Restplätze des Termin '" +
          termin +
          "' von " +
          restPlätze +
          " auf " +
          terminRest +
          " geändert!",
      );
    }
  }
}

function updateForm() {
  let termineHdrs = headers["Termine"];
  let termineRows = termineSheet.getLastRow() - 1; // first row = headers
  let termineCols = termineSheet.getLastColumn();
  let termineVals = termineSheet.getRange(2, 1, termineRows, termineCols).getValues();
  let termineNotes = termineSheet.getRange(2, 1, termineRows, 1).getNotes();
  // Logger.log("termine %s %s", termineVals.length, termineVals);
  let termineObjs = [];
  for (let i = 0; i < termineVals.length; i++) {
    if (!isEmpty(termineNotes[i][0])) continue;
    let terminObj: MapS2S = {};
    for (let hdr in termineHdrs) {
      let idx = termineHdrs[hdr];
      if (hdr == "Termin") {
        terminObj[hdr] = termineVals[i][idx - 1];
      } else {
        terminObj[hdr] = termineVals[i][idx - 1];
      }
    }
    let ok = true;
    // check if all cells of Termin row are nonempty
    for (let hdr in termineHdrs) {
      if (isEmpty(terminObj[hdr])) {
        Logger.log("In Termine Zeile mit leerem Feld %s", hdr);
        ok = false;
      }
      if (hdr == "Test") break;
    }
    if (ok) termineObjs.push(terminObj);
  }
  Logger.log("termineObjs=%s", termineObjs);

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formUrl = ss.getFormUrl();
  // Logger.log("formUrl2 %s", formUrl);
  let form: GoogleAppsScript.Forms.Form = FormApp.openByUrl(formUrl);
  let items = form.getItems();
  let termineItem: GoogleAppsScript.Forms.MultipleChoiceItem = null;
  for (let item of items) {
    //   let itemType = item.getType();
    //   Logger.log("title %s it %s %s", item.getTitle(), itemType, item.getIndex());
    if (item.getTitle() === terminFrage) {
      termineItem = item.asMultipleChoiceItem();
      break;
    }
  }
  if (termineItem == null) {
    SpreadsheetApp.getUi().alert('Das Formular hat keine Frage "' + terminFrage + '"!');
    return;
  }
  let choices = [];
  let descs = [];
  for (let terminObj of termineObjs) {
    let mr: string = terminObj["Termin"];

    let ok = true; // +terminObj["Rest"] > 0;
    let desc = mr + (ok ? ", freie Plätze: " + terminObj["Rest"] : ", einfach so kommen!");
    // Logger.log("desc %s", desc);
    descs.push(desc);
    if (ok) {
      let choice = termineItem.createChoice(mr);
      choices.push(choice);
    }
  }
  let beschreibung: string;
  if (choices.length === 0) {
    beschreibung = "Leider sind alle Termine vergeben!\n" + descs.join("\n");
    form.setAcceptingResponses(false);
    form.setCustomClosedFormMessage("Leider sind alle Termine vergeben!\n");
  } else {
    beschreibung =
      "Bitte einen Termin auswählen. Beachte die Anzahl noch freier Plätze!\n\n" + descs.join("\n");
    form.setAcceptingResponses(true);
    termineItem.setChoices(choices);
  }
  termineItem.setHelpText(beschreibung);
}

// I need any2str because a date copied to temp sheet showed as date.toString().
// A ' in front of the date came too late.
function any2Str(val: any): string {
  if (typeof val == "object" && "getUTCHours" in val) {
    return Utilities.formatDate(
      val,
      SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      "dd.MM.YYYY",
    );
  }
  return val.toString();
}

function printTerminMembers() {
  Logger.log("printTerminMembers");
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Termine") {
    SpreadsheetApp.getUi().alert("Bitte eine Zeile im Sheet 'Termine' selektieren");
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert("Bitte zuerst eine Zeile im Sheet 'Termine' selektieren");
    return;
  }
  let row = curCell.getRow();
  if (row < 2 || row > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Terminzeile selektieren",
    );
    return;
  }
  let rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  let rowNote = sheet.getRange(row, 1).getNote();
  if (!isEmpty(rowNote)) {
    SpreadsheetApp.getUi().alert("Die ausgewählte Zeile hat eine Notiz und ist deshalb ungültig");
    return;
  }
  let termin = rowValues[terminIndex - 1];
  let anmeldungenRows = anmeldungenSheet.getLastRow() - 1; // first row = headers
  let anmeldungenCols = anmeldungenSheet.getLastColumn();
  let anmeldungenVals: any[][];
  let anmeldungenNotes: string[][];
  // getRange with 0 rows throws an exception instead of returning an empty array
  if (anmeldungenRows < 1) {
    SpreadsheetApp.getUi().alert("Keine Anmeldungen gefunden");
    return;
  }
  let rows: string[][] = [];
  anmeldungenVals = anmeldungenSheet.getRange(2, 1, anmeldungenRows, anmeldungenCols).getValues();
  anmeldungenNotes = anmeldungenSheet.getRange(2, 1, anmeldungenRows, 1).getNotes();

  let bHdrs = headers["Anmeldungen"];
  // first row of temp sheet: the headers
  {
    let row: string[] = [];
    for (let [_, v] of printCols) {
      row.push(v);
    }
    rows.push(row);
  }
  for (let b = 0; b < anmeldungenRows; b++) {
    if (!isEmpty(anmeldungenNotes[b][0])) continue;
    let brow = anmeldungenVals[b];
    if (brow[anmeldeTerminIndex - 1] === termin) {
      let row: string[] = [];
      for (let [k, _] of printCols) {
        //for the ' see https://stackoverflow.com/questions/13758913/format-a-google-sheets-cell-in-plaintext-via-apps-script
        // otherwise, telefon number 089... is printed as 89
        let val = brow[bHdrs[k] - 1];
        row.push("'" + val);
      }
      rows.push(row);
    }
  }

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.insertSheet(termin);
  for (let row of rows) sheet.appendRow(row);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  let range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  sheet.setActiveSelection(range);
  printSelectedRange();
  // Utilities.sleep(10000);
  // ss.deleteSheet(sheet);
}

function objectToQueryString(obj: any) {
  return Object.keys(obj)
    .map(function (key) {
      return Utilities.formatString("&%s=%s", key, obj[key]);
    })
    .join("");
}

// see https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
let PRINT_OPTIONS = {
  size: 7, // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
  fzr: false, // repeat row headers
  portrait: true, // false=landscape
  fitw: true, // fit window or actual size
  gridlines: false, // show gridlines
  printtitle: true,
  sheetnames: true,
  pagenum: "UNDEFINED", // CENTER = show page numbers / UNDEFINED = do not show
  attachment: false,
};

let PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function printSelectedRange() {
  SpreadsheetApp.flush();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let range = sheet.getActiveRange();

  let gid = sheet.getSheetId();
  let printRange = objectToQueryString({
    c1: range.getColumn() - 1,
    r1: range.getRow() - 1,
    c2: range.getColumn() + range.getWidth() - 1,
    r2: range.getRow() + range.getHeight() - 1,
  });
  let url = ss.getUrl();
  Logger.log("url1 %s", url);
  let x = url.indexOf("/edit?");
  url = url.slice(0, x);
  url = url + "/export?format=pdf" + PDF_OPTS + printRange + "&gid=" + gid;
  Logger.log("url2 %s", url);
  let htmlTemplate = HtmlService.createTemplateFromFile("print.html");
  htmlTemplate.url = url;

  let ev = htmlTemplate.evaluate();

  SpreadsheetApp.getUi().showModalDialog(ev.setHeight(10).setWidth(100), "Drucke Auswahl");
}

function checkAnmeldungManually() {
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Anmeldungen") {
    SpreadsheetApp.getUi().alert("Bitte eine Zeile im Sheet 'Anmeldungen' selektieren");
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert("Bitte zuerst Teilnehmerzeile selektieren");
    return;
  }
  let rowIdx = curCell.getRow();
  if (rowIdx < 2 || rowIdx > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren",
    );
    return;
  }
  let rowNote = sheet.getRange(rowIdx, 1).getNote();
  if (!isEmpty(rowNote)) {
    SpreadsheetApp.getUi().alert("Die ausgewählte Zeile hat eine Notiz und ist deshalb ungültig");
    return;
  }
  let brange = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn());
  let brow = brange.getValues()[0];
  if (!isEmpty(brow[anmeldebestIndex - 1])) {
    SpreadsheetApp.getUi().alert("Die ausgewählte Buchung wurde schon bestätigt");
    return;
  }

  let e: SSEvent = {
    namedValues: {
      Vorname: [brow[vornameIndex - 1]],
      Nachname: [brow[nachnameIndex - 1]],
      "E-Mail-Adresse": [brow[mailIndex - 1]],
      [terminFrage]: [brow[anmeldeTerminIndex - 1]],
      Adresse: [brow[adresseIndex - 1]],
      Telefonnummer: [brow[telefonIndex - 1]],
    },
    range: brange,
  };
  checkAnmeldung(e);
}
