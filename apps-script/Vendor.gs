

function getVendorConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    VENDOR_WEBHOOK_URL: props.getProperty("VENDOR_WEBHOOK_URL") || "WEBHOOK_URL_HERE",
    WRITER_AUTH_TOKEN: props.getProperty("WRITER_AUTH_TOKEN") || "",

    VECTOR_STORE_ID: props.getProperty("VECTOR_STORE_ID") || "VECTOR_STORE_ID_HERE",
    TEMPLATE_VENDOR_ID: props.getProperty("TEMPLATE_VENDOR_ID") || "TEMPLATE_ID_HERE",
    PROPERTY_ROOT_FOLDER_ID: props.getProperty("PROPERTY_ROOT_FOLDER_ID") || "FOLDER_ID_HERE",

    MAX_OFFERS: parseInt(props.getProperty("MAX_OFFERS") || "5", 10),

    TIMEZONE: props.getProperty("TIMEZONE") || "Australia/Melbourne",
    BRAND_VOICE: (props.getProperty("BRAND_VOICE") || "assured,evidence-led,succinct")
      .split(",").map(s => s.trim()).filter(Boolean)
  };
}

/* =======================================================
   VENDOR REPORT AUTOMATION
   ======================================================= */
function runVendorAutomation_(e) {
  const CONFIG = getVendorConfig_();
  try {
    if (!e || !e.namedValues) throw new Error("No event payload (e.namedValues)");

    const inputs = mapFormToInputsVendor_(e, CONFIG);
    Logger.log("[Vendor] Inputs mapped (redacted).");

    const payload = {
      vectorStoreId: CONFIG.VECTOR_STORE_ID,
      inputs,
      brandVoice: CONFIG.BRAND_VOICE
    };

    const ai = callVendorWriter_(payload, CONFIG);

    const vendor = mergeIntoVendorDoc_({
      templateId: CONFIG.TEMPLATE_VENDOR_ID,
      title: makeDocTitleVendor_(inputs, CONFIG),
      inputs,
      ai,
      includeIssuesInVendor: shouldIncludeIssuesInVendor_(e)
    }, CONFIG);

    writeBackVendorLinks_(e, vendor);
    Logger.log("[Vendor] Generated doc + PDF.");

  } catch (err) {
    Logger.log(" Vendor automation error: " + err.message);
  }
}

/* ---------------------- Map Vendor Sheet + PropertyReport ---------------------- */
function mapFormToInputsVendor_(e, CONFIG) {
  const nv = e.namedValues || {};
  const v  = (k) => (nv[k] && nv[k][0]) ? String(nv[k][0]).trim() : "";
  const n  = (k) => parseFloat(v(k)) || 0;
  const d  = (k) => toISODateVendor_(v(k), CONFIG);
  const dt = (k) => toISODateTimeVendor_(v(k), CONFIG);
  const lines = (k) => splitLinesVendor_(v(k));
  const ids   = (k) => extractDriveIdsVendor_(v(k));

  const propertyId = v("Property ID");

  let property = {
    address: "", suburb: "", state: "", propertyType: "", saleOrLease: "",
    campaignWeek: v("Campaign Week"),
    periodFrom: d("Period From"),
    periodTo: d("Period To"),
    campaignMode: v("Campaign Mode"),
    agentNames: [],
    listingUrl: ""
  };

  // Pull core listing fields from PropertyReport
  try {
    const prSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PropertyReport");
    if (!prSheet) throw new Error("Sheet 'PropertyReport' not found.");

    const data = prSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idIndex = headers.findIndex(h => h.toLowerCase().replace(/\s+/g, "") === "propertyid");

    const match = data.find(r => String(r[idIndex]).trim() === String(propertyId).trim());
    if (match) {
      property.address      = match[headers.indexOf("Property Address")] || "";
      property.suburb       = match[headers.indexOf("Suburb")] || "";
      property.state        = match[headers.indexOf("State")] || "";
      property.propertyType = match[headers.indexOf("Property Type")] || "";
      property.saleOrLease  = match[headers.indexOf("Sale or Lease")] || "";
      property.agentNames   = [
        match[headers.indexOf("Agent 1 Name")] || "",
        match[headers.indexOf("Agent 2 Name")] || ""
      ].filter(Boolean);
      property.listingUrl   = match[headers.indexOf("Listing URL (if live)")] || "";
    }
  } catch (err) {
    Logger.log(" PropertyReport lookup error: " + err.message);
  }

  return {
    propertyId,
    property,
    kpis: {
      portals: {
        reaViews: n("REA Views"),
        domainViews: n("Domain Views"),
        enquiries: n("Enquiries"),
        saves: n("Saves/Shortlists")
      },
      inspections: {
        opensHeld: n("Opens Held"),
        attendees: n("Attendees"),
        privateInspections: n("Private Inspections")
      },
      marketing: {
        emailOpens: n("Email Opens"),
        emailClicks: n("Email Clicks"),
        socialReach: n("Social Reach"),
        spendThisWeek: n("Spend This Week"),
        spendToDate: n("Spend To Date"),
        budgetRemaining: n("Budget Remaining")
      }
    },
    feedback: {
      buyerQuotes: lines("Buyer Quotes"),
      priceGuidance: v("Price Guidance")
    },
    offers: collectOffersVendor_(nv, CONFIG.MAX_OFFERS),
    risksIssues: lines("Risks/Issues"),
    nextActions: lines("Next Actions"),
    nextOpenHome: dt("Next Open Home"),
    photos: {
      property: { fileIds: ids("Property Photos"), captions: lines("Property Photo Captions") },
      issues:   { fileIds: ids("Issues Photos"), captions: lines("Issues Photo Captions") }
    },
    includeIssues: shouldIncludeIssuesInVendor_(e),
    agentNotes: v("Agent Notes")
  };
}

/* ---------------------- Generate Docs ---------------------- */
function mergeIntoVendorDoc_(args, CONFIG) {
  const { templateId, title, inputs, ai, includeIssuesInVendor } = args;

  const parent = DriveApp.getFolderById(CONFIG.PROPERTY_ROOT_FOLDER_ID);

  // Avoid leaking real addresses in public demos; keep folder naming generic if needed
  const safeName = (inputs.property.address || "Property").replace(/[\\/:*?"<>|]+/g, " ");
  const folderName = `Vendor Report – ${safeName} – ${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HHmm")}`;

  const folder = parent.createFolder(folderName);
  const template = DriveApp.getFileById(templateId);

  const copy = template.makeCopy(title, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // Merge AI fields ({{key}} placeholders)
  Object.keys(ai || {}).forEach(k => {
    const val = String(ai[k] || "");
    body.replaceText(escapeFindVendor_(`{{${k}}}`), val);
    body.replaceText(escapeFindVendor_(`{{${k.replace(/[^\w]+/g, "_")}}}`), val);
  });

  // Merge fixed fields
  body.replaceText(escapeFindVendor_("{{PropertyID}}"), inputs.propertyId || "");
  body.replaceText(escapeFindVendor_("{{Property.Address}}"), inputs.property.address || "");
  body.replaceText(escapeFindVendor_("{{Property.Type}}"), inputs.property.propertyType || "");
  body.replaceText(escapeFindVendor_("{{Property.CampaignWeek}}"), inputs.property.campaignWeek || "");
  body.replaceText(escapeFindVendor_("{{Property.PeriodFrom}}"), inputs.property.periodFrom || "");
  body.replaceText(escapeFindVendor_("{{Property.PeriodTo}}"), inputs.property.periodTo || "");
  body.replaceText(escapeFindVendor_("{{Method of Sale}}"), inputs.property.campaignMode || "");

  fillReportMetaVendor_(doc, inputs, CONFIG);
  insertGalleryVendor_(doc, "{{IMG_PROPERTY_GALLERY}}", inputs.photos.property.fileIds, inputs.photos.property.captions, 480);

  if (includeIssuesInVendor && (inputs.photos.issues.fileIds || []).length) {
    insertGalleryVendor_(doc, "{{IMG_ISSUES_GALLERY}}", inputs.photos.issues.fileIds, inputs.photos.issues.captions, 480);
  } else {
    body.replaceText(escapeFindVendor_("{{IMG_ISSUES_GALLERY}}"), "");
  }

  doc.saveAndClose();

  const pdf = folder.createFile(copy.getAs("application/pdf").setName(title + ".pdf"));
  return { docUrl: copy.getUrl(), pdfUrl: pdf.getUrl() };
}

/* ---------------------- AI Writer Call ---------------------- */
function callVendorWriter_(payload, CONFIG) {
  const endpoint = CONFIG.VENDOR_WEBHOOK_URL;
  const headers = { "Content-Type": "application/json" };
  if (CONFIG.WRITER_AUTH_TOKEN) headers["Authorization"] = "Bearer " + CONFIG.WRITER_AUTH_TOKEN;

  const options = {
    method: "post",
    headers,
    payload: JSON.stringify(payload),
    followRedirects: true,
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(endpoint, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText();

  Logger.log(`[Vendor AI] HTTP ${code} (response truncated)`);
  if (code >= 300) throw new Error(`Vendor AI error ${code}: ${text.slice(0, 300)}`);

  const json = JSON.parse(text || "{}");
  if (!json.ai) throw new Error("No AI payload returned from Vendor pipeline.");
  return json.ai;
}

/* ---------------------- Helpers ---------------------- */
function escapeFindVendor_(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function toISODateVendor_(s,CONFIG){ if(!s) return ''; const d=new Date(s); return isNaN(d)?'':Utilities.formatDate(d,CONFIG.TIMEZONE,'yyyy-MM-dd'); }
function toISODateTimeVendor_(s,CONFIG){ if(!s) return ''; const d=new Date(s); return isNaN(d)?'':Utilities.formatDate(d,CONFIG.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"); }
function splitLinesVendor_(s){ return String(s||'').split(/\r?\n/).map(x=>x.trim()).filter(Boolean); }
function extractDriveIdsVendor_(cell){
  if(!cell) return [];
  const parts = String(cell).split(',').map(x=>x.trim());
  const ids=[];
  parts.forEach(u=>{
    let m = u.match(/\/d\/([a-zA-Z0-9_-]+)/) || u.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if(m) ids.push(m[1]);
  });
  return ids;
}
function shouldIncludeIssuesInVendor_(e){
  const nv=e.namedValues||{};
  const ans=(nv['Include Issues Photos in Vendor Copy?'] && nv['Include Issues Photos in Vendor Copy?'][0]) || '';
  return String(ans).trim().toLowerCase()==='yes';
}
function fillReportMetaVendor_(doc, inputs, CONFIG){
  const body=doc.getBody();
  const when = inputs?.property?.periodTo ? new Date(inputs.property.periodTo) : new Date();
  const reportDate = Utilities.formatDate(when, CONFIG.TIMEZONE, 'd MMM yyyy');
  const preparedBy = (inputs?.property?.agentNames||[]).join(', ') || 'Agent';
  body.replaceText(escapeFindVendor_('{{Report.Date}}'), reportDate);
  body.replaceText(escapeFindVendor_('{{Report.PreparedBy}}'), preparedBy);
}
function insertGalleryVendor_(doc, ph, ids, caps, w){
  const body=doc.getBody();
  const range=body.findText(escapeFindVendor_(ph));
  if(!range) return;

  const el=range.getElement().asText();
  el.deleteText(range.getStartOffset(), range.getEndOffsetInclusive());
  const p=el.getParent().asParagraph();

  (ids||[]).forEach((id,i)=>{
    try{
      const blob=DriveApp.getFileById(id).getBlob();
      const img=p.appendInlineImage(blob);
      if(w) img.setWidth(w);
      const cap=(caps && caps[i]) ? caps[i] : '';
      if(cap) p.appendText('\n'+cap);
      if(i<ids.length-1) body.appendParagraph('');
    }catch(err){
      p.appendText(`\n[Image unavailable]`);
    }
  });
}
function writeBackVendorLinks_(e, links){
  const sh=e.range.getSheet();
  const row=e.range.getRow();
  const c1=ensureHeaderVendor_(sh,'Generated Vendor Doc URL');
  const c2=ensureHeaderVendor_(sh,'Generated Vendor PDF URL');
  sh.getRange(row,c1).setValue(links.docUrl||'');
  sh.getRange(row,c2).setValue(links.pdfUrl||'');
}
function ensureHeaderVendor_(sh, h){
  const lastCol=sh.getLastColumn();
  const headers=lastCol ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
  let i=headers.indexOf(h);
  if(i!==-1) return i+1;
  sh.getRange(1,lastCol+1).setValue(h);
  return lastCol+1;
}
function collectOffersVendor_(nv, maxN){
  const offers=[];
  for(let i=1;i<=maxN;i++){
    const buyer=(nv[`Offer ${i} Buyer`]||[''])[0];
    if(!buyer) continue;
    offers.push({
      buyer,
      status:(nv[`Offer ${i} Status`]||[''])[0],
      amount:(nv[`Offer ${i} Amount (AUD)`]||[''])[0],
      conditions:(nv[`Offer ${i} Conditions`]||[''])[0]
    });
  }
  return offers;
}
function makeDocTitleVendor_(inputs, CONFIG){
  const name = inputs.property.address || 'Property';
  return `Vendor Report - ${name} - ${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HHmm')}`;
}
