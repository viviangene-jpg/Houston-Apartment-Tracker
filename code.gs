const SHEET_URL = "https://docs.google.com/spreadsheets/d/14j7JJbnia5nlLhrF8grfooid8sMS0q_9kRQOczeaies/edit";
const SHEET_NAME = "tracker"; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HTX Apartment Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 1-to-1 SAVE MAPPING (A-X)
 * Strictly matches your provided string order.
 */
function handleSaveRequest(data) {
  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  const rowValues = [[
    data.name || "",           // A [0] Name
    data.address || "",        // B [1] Address
    data.zip || "",            // C [2] Zip
    data.units || "",          // D [3] Units
    data.status || "",         // E [4] Status
    data.class || "",          // F [5] Class
    data.price || "",          // G [6] Price
    data.foreclosure || "",    // H [7] Foreclosure
    data.priceperunit || "",   // I [8] PricePerUnit
    data.askingcap || "",      // J [9] AskingCap
    data.income || "",         // K [10] Income
    data.expense || "",        // L [11] Expense
    data.expratio || "",       // M [12] ExpRatio
    data.noi || "",            // N [13] NOI
    data.targetcap || "",      // O [14] TargetCap
    data.targetoffer || "",    // P [15] targetOffer
    data.downpct || "",        // Q [16] DownPct
    data.downamt || "",        // R [17] DownAmt
    data.interestrate || "",   // S [18] InterestRate
    data.amortization || "",   // T [19] Amortization
    data.coc || "",            // U [20] COC
    data.dscr || "",           // V [21] DSCR
    data.notes || "",          // W [22] Notes
    data.timestamp || new Date().toLocaleString() // X [23] Timestamp
  ]];

  var rowToUpdate = (data.rowid && data.rowid !== "") ? parseInt(data.rowid) : sheet.getLastRow() + 1;
  sheet.getRange(rowToUpdate, 1, 1, 24).setValues(rowValues);
  return "Success";
}

/**
 * 1-to-1 PULL MAPPING (A-X)
 */
function getInventoryJSON() {
  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    results.push({
      rowid: i + 1,
      name: row[0],         address: row[1],      zip: row[2],          units: row[3],
      status: row[4],       class: row[5],        price: row[6],        foreclosure: row[7],
      priceperunit: row[8], askingcap: row[9],    income: row[10],      expense: row[11],
      expratio: row[12],    noi: row[13],         targetcap: row[14],   targetoffer: row[15],
      downpct: row[16],     downamt: row[17],     interestrate: row[18], amortization: row[19],
      coc: row[20],         dscr: row[21],        notes: row[22],       timestamp: row[23]
    });
  }
  return JSON.stringify(results);
}
