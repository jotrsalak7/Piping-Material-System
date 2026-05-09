/**
 * Google Apps Script for Piping Material Intelligence
 * Serves data from "Monitoring-Pi-Matl." sheet as JSON via doGet()
 */

function doGet(e) {
  try {
    const sheetName = 'Monitoring-Pi-Matl.';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: `Sheet "${sheetName}" not found` })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get all data
    const data = sheet.getDataRange().getValues();
    
    // First row is headers
    const headers = data[0];
    const rows = data.slice(1);
    
    // Convert to array of arrays (compact format)
    const compactData = rows.map(row => row);
    
    const response = {
      success: true,
      timestamp: new Date().toISOString(),
      count: rows.length,
      data: compactData
    };
    
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * DEPLOYMENT INSTRUCTIONS:
 * 
 * 1. Create a new Google Apps Script project at script.google.com
 * 2. Paste this code into the editor
 * 3. Click "Deploy" → "New deployment" → Select "Web app"
 * 4. Execute as: Your email
 *    Who has access: Anyone
 * 5. Copy the deployment URL (looks like: https://script.google.com/macros/d/DEPLOYMENT_ID/usercss)
 * 6. Paste this URL into the HTML file where it says:
 *    const APPS_SCRIPT_URL = 'YOUR_URL_HERE';
 * 
 * SHEET STRUCTURE:
 * The "Monitoring-Pi-Matl." sheet should have columns in this order:
 * 
 * A: area
 * B: package
 * C: bomTitle
 * D: identCode
 * E: partName
 * F: desc
 * G: lSize
 * H: qtyReq
 * I: unit
 * J: poNo
 * K: vendor
 * L: statusPO (e.g., "WAITING PO", "ISSUED", "CONFIRMED")
 * M: mfgPct (manufacturing percentage)
 * N: eta (estimated arrival)
 * O: ata (actual arrival)
 * P: pctDelv (percent delivered)
 * Q: qtySite (quantity on site)
 * R: shortage (negative = shortage, positive = surplus)
 * S: statusMtl (material status: "On Site", "Partial Delv.", "-")
 * 
 * Start data from row 2 (row 1 should be headers).
 */
