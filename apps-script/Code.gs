
function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    // Marketing
    SLIDES_TEMPLATE_ID: props.getProperty("SLIDES_TEMPLATE_ID") || "TEMPLATE_ID_HERE",
    FLYER_TEMPLATE_ID: props.getProperty("FLYER_TEMPLATE_ID") || "TEMPLATE_ID_HERE",
    SOCIAL_TEMPLATE_ID: props.getProperty("SOCIAL_TEMPLATE_ID") || "TEMPLATE_ID_HERE",
    OUTPUT_FOLDER_ID: props.getProperty("OUTPUT_FOLDER_ID") || "FOLDER_ID_HERE",
    MARKETING_WEBHOOK_URL: props.getProperty("MARKETING_WEBHOOK_URL") || "WEBHOOK_URL_HERE",

    
    VENDOR_FORM_BASE: props.getProperty("VENDOR_FORM_BASE") || "https://docs.google.com/forms/d/e/FORM_ID/viewform?usp=pp_url",
    VENDOR_ENTRY_ID: props.getProperty("VENDOR_ENTRY_ID") || "ENTRY_ID_HERE",

    // Common
    TIMEZONE: props.getProperty("TIMEZONE") || "Australia/Melbourne"
  };
}


function doGet(e) {
  const CONFIG = getConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PropertyReport");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Sheet 'PropertyReport' not found." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift().map(h => String(h).trim());

  const json = data.map(r => {
    const obj = Object.fromEntries(headers.map((h, i) => [h, r[i]]));
    if (obj.PropertyID) {
      obj.VendorReportFormURL =
        `${CONFIG.VENDOR_FORM_BASE}&entry.${CONFIG.VENDOR_ENTRY_ID}=${encodeURIComponent(obj.PropertyID)}`;
    }
    return obj;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ data: json }))
    .setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PropertyReport");
    if (!sheet) throw new Error("Sheet 'PropertyReport' not found.");

    const CONFIG = getConfig_();
    const body = JSON.parse(e.postData.contents || "{}");
    const propertyId = body.PropertyID || Utilities.getUuid();

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const newRow = headers.map(h => {
      switch (h) {
        case "Timestamp": return new Date();
        case "PropertyID": return propertyId;
        case "Client Name": return body.name || "";
        case "Email": return body.email || "";
        case "Client Phone Number": return body.phone || "";
        case "Client Address": return body.address || "";
        case "Client Company": return body.notes || "";
        default: return "";
      }
    });

    sheet.appendRow(newRow);

    const vendorURL =
      `${CONFIG.VENDOR_FORM_BASE}&entry.${CONFIG.VENDOR_ENTRY_ID}=${encodeURIComponent(propertyId)}`;

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, propertyId, vendorURL })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}


function onFormSubmit(e) {
  try {
    if (!e || !e.range) throw new Error("Missing event object (e).");

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const propertyIdCol =
      headers.findIndex(h => String(h).toLowerCase().replace(/\s+/g, "") === "propertyid") + 1;

    if (propertyIdCol <= 0) {
      Logger.log(" No 'Property ID' column found.");
      return;
    }

    const row = e.range.getRow();
    const currentId = sheet.getRange(row, propertyIdCol).getValue();
    if (!currentId) sheet.getRange(row, propertyIdCol).setValue(Utilities.getUuid());

    Logger.log("🧾 Sheet detected: " + sheetName);

    if (sheetName === "MarketingPack") {
      Logger.log("Triggering marketing automation...");
      runMarketingAutomation_(e);
    } else if (sheetName === "VendorReport") {
      Logger.log("Triggering vendor automation...");
      // Implemented in Vendor.gs
      runVendorAutomation_(e);
    } else {
      Logger.log("Skipped automation: " + sheetName);
    }

  } catch (err) {
    Logger.log("onFormSubmit error: " + err.message);
  }
}


function runMarketingAutomation_(e) {
  try {
    Logger.log(" Marketing automation started");
    if (!e || !e.namedValues) throw new Error("Missing event payload (namedValues).");

    const inputs = readInputs_(e.namedValues);
    // const ai = callAiContentGenerator_(inputs);
    // const pack = ensureListingFolder_(inputs);

    // const outSign = buildSlides_(inputs, ai, pack.folder);
    // writeBackLink_(e, 'Generated Signboard Slides URL', outSign.url);

    // ... flyer, tile, caption, folder link

    Logger.log(" Marketing automation finished (demo-safe skeleton).");
  } catch (err) {
    Logger.log(" Marketing automation error: " + err.message);
  }
}


function readInputs_(nv) {
  const propertyId = getNVByPrefix_(nv, "Property ID");
  const aiNotes = getNVByPrefix_(nv, "What makes this property stand out");
  const targetProfile = getNVByPrefix_(nv, "Target buyer");

  
  const heroId = readImageIdByPrefix_(nv, "Hero Image");
  const thumb1Id = readImageIdByPrefix_(nv, "Additional Image 1");
  const thumb2Id = readImageIdByPrefix_(nv, "Additional Image 2");
  const thumb3Id = readImageIdByPrefix_(nv, "Floor Plan");

 
  let address = "", suburb = "", state = "", propertyType = "", saleOrLease = "", listingUrl = "";
  let buildingArea = "", siteArea = "";
  let agent1Name = "", agent1Mobile = "", agent2Name = "", agent2Mobile = "";

  try {
    const prSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PropertyReport");
    if (!prSheet) throw new Error("Sheet 'PropertyReport' not found.");

    const data = prSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idIndex = headers.findIndex(h => h.toLowerCase().replace(/\s+/g, "") === "propertyid");

    const match = data.find(r => String(r[idIndex]).trim() === String(propertyId).trim());
    if (match) {
      address = match[headers.indexOf("Property Address")] || "";
      suburb = match[headers.indexOf("Suburb")] || "";
      state = match[headers.indexOf("State")] || "";
      propertyType = match[headers.indexOf("Property Type")] || "";
      saleOrLease = match[headers.indexOf("Sale or Lease")] || "";
      listingUrl = match[headers.indexOf("Listing URL (if live)")] || "";
      buildingArea = match[headers.indexOf("Building Area (m²)")] || "";
      siteArea = match[headers.indexOf("Site Area / Land (m²)")] || "";
      agent1Name = match[headers.indexOf("Agent 1 Name")] || "";
      agent1Mobile = match[headers.indexOf("Agent 1 Mobile")] || "";
      agent2Name = match[headers.indexOf("Agent 2 Name")] || "";
      agent2Mobile = match[headers.indexOf("Agent 2 Mobile")] || "";
    }
  } catch (err) {
    Logger.log(" PropertyReport lookup error: " + err.message);
  }

  return {
    propertyId, address, suburb, state, propertyType, saleOrLease,
    listingUrl, buildingArea, siteArea,
    aiNotes, targetProfile,
    heroId, thumb1Id, thumb2Id, thumb3Id,
    agent1Name, agent1Mobile, agent2Name, agent2Mobile
  };
}


function getNVByPrefix_(nv, prefix) {
  const key = Object.keys(nv || {}).find(k => String(k).trim().startsWith(prefix));
  return key ? String(nv[key][0] || "").trim() : "";
}

function readImageIdByPrefix_(nv, prefix) {
  const val = getNVByPrefix_(nv, prefix);
  if (!val) return "";
  // Support patterns: /d/<id> or id=<id>
  const m = String(val).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(val).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : "";
}
