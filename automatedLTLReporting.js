const odDataSheet = SpreadsheetApp.getActive().getSheetByName("ODFL Data");
const nsDataSheet = SpreadsheetApp.getActive().getSheetByName("NS Data");
const datasetSheet = SpreadsheetApp.getActive().getSheetByName("Dataset");
const dcInfoSheet = SpreadsheetApp.getActive().getSheetByName("DC Info");

// how to handle reports with no data? - unlikely for OD but possible
function getODEmailReport() {
  const gmailThread = GmailApp.search("label:odfl-ship-report", 0, 1)[0];
  const csvAttachment = gmailThread
    .getMessages()
    [gmailThread.getMessageCount() - 1].getAttachments();
  const csvFile = Utilities.parseCsv(csvAttachment[0].getDataAsString());
  return csvFile;
}

function cleanOldDominionData() {
  const sheet = odDataSheet;
  const data = sheet.getDataRange().getValues();
  // aka no headers
  let values = data.slice(1);

  for (let i = 0; i < values.length; i++) {
    //  remove '#' from PO # column
    values[i][10] = String(values[i][10].replace("#", ""));
  }
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
}

function updateOldDominionData() {
  const csv = getODEmailReport();
  odDataSheet.getDataRange().clearContent();
  odDataSheet.getRange(1, 1, csv.length, csv[0].length).setValues(csv);
  cleanOldDominionData();
}

// Will need to update gmail labeling once the new report comes in tomorrow 7/20 (I changed the subject line)
// remember to use memo field to store the original due date for orders until the native NS field can be enabled
function getNSEmailReport() {
  const gmailThread = GmailApp.search("label:ns-ltl-report ", 0, 1)[0];
  const csvAttachment = gmailThread
    .getMessages()
    [gmailThread.getMessageCount() - 1].getAttachments();
  const csvFile = Utilities.parseCsv(csvAttachment[0].getDataAsString());
  return csvFile;
}

// added a silent failure if NS CSV data has only header row here
// might be wise to figure out a better way of error handling
function cleanNetsuiteData() {
  const sheet = nsDataSheet;
  const data = sheet.getDataRange().getValues();
  // Logger.log(data.length);
  if (data.length !== 1) {
    let values = data.slice(1);
    let sortRange = sheet.getRange(
      2,
      1,
      sheet.getLastRow(),
      sheet.getLastColumn()
    );

    // might be a way to keep original cell value in vlookup but for now printing an error message will stopgap
    for (let i = 0; i < values.length; i++) {
      const vlookup = `=IFERROR(VLOOKUP(A${[
        i + 2,
      ]},'DC Info'!$A$2:$B,2,FALSE),"Interal ID not found in DC Info sheet")`;
      values[i][1] = vlookup;
    }

    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    // sorts NS Data by Fulfillment Date
    sortRange.sort(4);
  }
}

function addNetsuiteCSVToSheet() {
  const csv = getNSEmailReport();
  nsDataSheet.getDataRange().clearContent();
  nsDataSheet.getRange(1, 1, csv.length, csv[0].length).setValues(csv);
  cleanNetsuiteData();
}

// might be able to clean up variable declaration some (removing filter funcs)
// may also need to adjust where columns are located on Dataset sheet
function updateNetsuiteData() {
  // added this form to bundle all NS functions into one call
  addNetsuiteCSVToSheet();
  const dataSheet = datasetSheet;
  const nsData = nsDataSheet
    .getRange(2, 1, nsDataSheet.getLastRow(), nsDataSheet.getLastColumn())
    .getValues()
    .filter(function (row) {
      return row.filter(Boolean).length > 0;
    });
  let dataset = datasetSheet
    .getRange(2, 1, datasetSheet.getLastRow(), nsDataSheet.getLastColumn())
    .getValues()
    .filter(function (row) {
      return row.filter(Boolean).length > 0;
    });

  for (let i in nsData) {
    dataset.push(nsData[i]);
  }
  dataSheet
    .getRange(2, 1, dataset.length, dataset[0].length)
    .setValues(dataset);
}

// adds sheets formulas into blank shipping details columns to pull data from ODFL Data sheet
function updateShipmentInfo() {
  const datasetRange = datasetSheet
    .getRange(2, 1, datasetSheet.getLastRow(), 11)
    .getValues();
  const formulaRange = datasetSheet
    .getRange(2, 8, datasetRange.length, 4)
    .getValues();
  let filledCellCount = 0;

  for (let i = datasetRange.length - 1; i >= 0; i--) {
    // the i+2 has something to do with where i is as it relates to rows... haven't figured out a better way to solve that problem yet
    // it works though!
    const actualShipDate = `=IFNA(INDEX(ActualPickupDate_OD,MATCH(C${
      i + 2
    },POnumber_OD,0)),"")`;
    const arrivedODYard = `=IFNA(INDEX(ArrivalDate_OD,MATCH(C${
      i + 2
    },POnumber_OD,0)),"")`;
    const deliveryDate = `=IFNA(INDEX(DeliveryDate_OD,MATCH(C${
      i + 2
    },POnumber_OD,0)),"")`;
    const proNumber = `=IFNA(INDEX(PROnumber_OD,MATCH(C${
      i + 2
    },POnumber_OD,0)),"")`;
    // Logger.log(formulaRange[i])
    // use this to break the loop if it gets high enough
    // probably will check to see if OD delivery column values are filled or not
    // this is just to keep the loop from having to run all the way to the end every time

    if (formulaRange[i][0] === "") {
      formulaRange[i][0] = actualShipDate;
    }
    if (formulaRange[i][1] === "") {
      formulaRange[i][1] = arrivedODYard;
    }
    if (formulaRange[i][2] === "") {
      formulaRange[i][2] = deliveryDate;
    }
    if (formulaRange[i][3] === "") {
      formulaRange[i][3] = proNumber;
    }
    if (formulaRange[i][2] != "") filledCellCount++;
    if (filledCellCount > 30) break;
  }
  datasetSheet.getRange(2, 8, datasetRange.length, 4).setValues(formulaRange);
}

// autofills the 6 columns of sheets formulas on the right hand side of the dataset sheet.
// # of columns is hardcoded at the moment, if formulas are added this will need to be updated
// currently filling more rows than I'd like - extends several rows past
function fillRightHandFormulas() {
  const firstRow = datasetSheet.getRange(["L2:Q2"]);
  const formulaRows = datasetSheet.getRange(
    2,
    12,
    datasetSheet.getLastRow(),
    6
  );
  firstRow.autoFill(formulaRows, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

// Get rid of formulas from Netsuite and Shipping details section while keeping values
// ie. locks in the data for first 11 columns without any pesky formulas sticking around
// will be the last function called
function pasteValsOnlyEquiv() {
  const rngCopyValsOnly = datasetSheet
    .getRange(2, 1, datasetSheet.getLastRow(), 11)
    .getValues();
  datasetSheet
    .getRange(2, 1, datasetSheet.getLastRow(), 11)
    .setValues(rngCopyValsOnly);
}

// Added
function autoUpdateDataset() {
  // Netsuite functions will not execute if there is no data - see cleanNSData for details
  updateNetsuiteData();
  updateOldDominionData();
  updateShipmentInfo();
  // this fills lots of extra rows - how to slim it down some?
  fillRightHandFormulas();
  pasteValsOnlyEquiv();
}

// autoruns script every morning at 9am
ScriptApp.newTrigger(autoUpdateDataset)
  .timeBased()
  .atHour(9)
  .everyDays(1)
  .inTimezone("America/New_York")
  .create();
