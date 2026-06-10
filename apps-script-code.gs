const SHEET_NAME = "RNA Responses";
const NOTIFY_EMAIL = "naycoyap@goodneighbors.ph";
const LOCKED_DRIVE_UPLOAD_URL = "https://drive.google.com/drive/folders/1Ex3kuycGsDj973oVse4Il_HbGd9M6Xuj?usp=sharing";

function doPost(e) {
  const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  const headers = body.headers || [];
  const values = body.values || [];
  const raw = body.raw || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0 && headers.length) {
    sheet.appendRow(headers);
  }

  if (values.length) {
    sheet.appendRow(values);
  }

  const reportName = buildReportFileName(raw);
  const reportHtml = buildReportHtml(raw);
  const reportBlob = Utilities.newBlob(reportHtml, "application/msword", reportName);
  const reportFile = DriveApp.createFile(reportBlob);
  const recipients = reportRecipients(raw);
  recipients.forEach((email) => reportFile.addViewer(email));

  MailApp.sendEmail({
    to: recipients.join(","),
    subject: `Rapid Needs Assessment Submitted: ${raw.emergencyName || "Unnamed Emergency"}`,
    body: `A Rapid Needs Assessment submission was received.\n\nThe formatted Word report is attached and shared through Google Drive.\n\nEmergency: ${raw.emergencyName || "Not supplied"}\nAssessment period: ${raw.assessmentFrom || "Not supplied"} to ${raw.assessmentTo || "Not supplied"}\nLocation: ${[raw.regions, raw.provinces, raw.cities, raw.barangays].filter(Boolean).join(" / ") || "Not supplied"}\nReport file: ${reportFile.getUrl()}`,
    htmlBody: `
      <p>A Rapid Needs Assessment submission was received.</p>
      <p>The formatted Word report is attached and shared through Google Drive.</p>
      <p><strong>Emergency:</strong> ${escapeHtml(raw.emergencyName || "Not supplied")}</p>
      <p><strong>Assessment period:</strong> ${escapeHtml(raw.assessmentFrom || "Not supplied")} to ${escapeHtml(raw.assessmentTo || "Not supplied")}</p>
      <p><strong>Location:</strong> ${escapeHtml([raw.regions, raw.provinces, raw.cities, raw.barangays].filter(Boolean).join(" / ") || "Not supplied")}</p>
      <p><strong>Word report:</strong> <a href="${reportFile.getUrl()}">Open shared report</a></p>
    `,
    attachments: [reportBlob]
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, reportUrl: reportFile.getUrl() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildReportFileName(raw) {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const name = String(raw.emergencyName || "rapid-needs-assessment")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 70) || "rapid-needs-assessment";
  return `${name}-${date}.doc`;
}

function buildReportHtml(raw) {
  const locations = parseJson(raw.locationsJson, []);
  const sectionB = parseJson(raw.sectionBJson, {});
  const sectionC = parseJson(raw.sectionCJson, []);
  const sectionD = parseJson(raw.sectionDJson, []);
  const members = parseJson(raw.sectionEJson, []);
  const lguContacts = parseJson(raw.sectionELguContactsJson, []);

  const locationRows = locations.map((loc) => [
    loc.region,
    loc.province,
    loc.city,
    loc.barangay,
    loc.affectedHouseholds || 0,
    loc.affectedIndividuals || 0,
    loc.damagedHealthFacilities || 0,
    loc.damagedSchools || 0,
    loc.affectedLearners || 0,
    loc.damagedL3WaterSystems || 0
  ]);

  const populationRows = [];
  Object.keys(sectionB || {}).forEach((locationId) => {
    const detail = sectionB[locationId] || {};
    Object.keys(detail.ageGroups || {}).forEach((key) => {
      const row = detail.ageGroups[key] || {};
      const female = Number(row.female || 0);
      const male = Number(row.male || 0);
      if (female || male) populationRows.push([locationId, key, female, male, female + male]);
    });
  });

  const needsRows = sectionC.map((item) => [
    item.region,
    item.province,
    item.municipality,
    item.barangay || "",
    item.exposureLevel || "",
    (item.exposureIndicators || []).join("; "),
    item.sector || "",
    item.gradeRiskLevel || "",
    item.overallRiskProfile || "",
    (item.observations || []).join("; ") || item.observation || "",
    (item.priorityNeeds || []).join("; "),
    item.additionalInformation || ""
  ]);

  const responseRows = sectionD.map((item) => [
    item.region,
    item.province,
    item.municipality,
    item.barangay || "",
    item.responseType || "",
    item.activity || item.localResponse || "",
    (item.actors || []).join("; ") || item.actor || "",
    item.details || item.additionalInformation || ""
  ]);

  const memberRows = members.map((item) => [
    item.name,
    item.role,
    item.organization,
    item.email,
    item.mobile
  ]);

  const lguRows = lguContacts.map((item) => [
    item.name,
    item.designation,
    item.office,
    item.mobile,
    item.email
  ]);

  return `<!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          @page { size: landscape; margin: .45in; }
          body { font-family: Arial, sans-serif; color: #222; font-size: 9pt; line-height: 1.35; }
          h1 { color: #5f772b; font-size: 20pt; margin: 0 0 4px; }
          h2 { color: #5f772b; font-size: 13pt; margin: 18px 0 8px; border-bottom: 1px solid #b7c996; padding-bottom: 4px; }
          h3 { font-size: 10.5pt; margin: 12px 0 6px; color: #2f3d33; }
          p { margin: 5px 0; }
          .report-header { border-bottom: 3px solid #8aa540; margin-bottom: 14px; padding-bottom: 10px; }
          .report-subtitle { font-size: 11pt; color: #4b5750; margin: 0; }
          .meta { width: 100%; border-collapse: collapse; margin: 8px 0 14px; table-layout: fixed; }
          .meta td { border: 0; padding: 3px 6px; font-size: 8.5pt; }
          .meta td:first-child, .kv td:first-child { font-weight: bold; color: #5f772b; width: 22%; }
          table { border-collapse: collapse; width: 100%; table-layout: fixed; margin: 8px 0 12px; }
          th, td { border: 1px solid #999; padding: 4px; text-align: left; vertical-align: top; word-wrap: break-word; font-size: 7.5pt; }
          th { background: #eef4e8; color: #26362b; text-transform: uppercase; }
          .kv td { font-size: 8.5pt; }
          .wide th, .wide td { font-size: 6.8pt; padding: 3px; }
          .narrative { white-space: pre-wrap; text-align: justify; }
          .page-break { page-break-before: always; }
        </style>
      </head>
      <body>
        <div class="report-header">
          <h1>Rapid Needs Assessment Form</h1>
          <p class="report-subtitle">${escapeHtml(raw.emergencyName || "Rapid Needs Assessment Report")}</p>
        </div>
        <table class="meta"><tbody>
          <tr><td>Generated</td><td>${escapeHtml(new Date())}</td><td>Assessment Period</td><td>${escapeHtml(raw.assessmentFrom || "")} to ${escapeHtml(raw.assessmentTo || "")}</td></tr>
          <tr><td>Date of Event</td><td>${escapeHtml(raw.eventDate || "")}</td><td>Primary Locations</td><td>${escapeHtml(locations.map((loc) => loc.city || loc.province || loc.region).filter(Boolean).slice(0, 8).join("; "))}</td></tr>
        </tbody></table>
        <h2>Section A: Emergency Details</h2>
        ${keyValueTable([
          ["Type of Hazard", raw.hazardType],
          ["Sub Type", raw.hazardSubType],
          ["Detailed Hazard", raw.detailedHazard],
          ["Name of Emergency", raw.emergencyName],
          ["Date of Event", raw.eventDate],
          ["Description of Emergency", raw.emergencyDescription]
        ])}
        <h3>Table of Affected Population</h3>
        ${tableHtml(["Region", "Province", "Municipality/City", "Barangay", "No. of HH Affected", "No. of Individuals", "Damaged Health Facilities", "Damaged Schools", "Affected Learners", "Damaged L3 Water Systems"], locationRows, "wide")}
        <h2>Section B: Affected Population</h2>
        ${tableHtml(["Location ID", "Age Group", "Female", "Male", "Total"], populationRows, "wide")}
        <div class="page-break"></div>
        <h2>Section C: Needs Assessment</h2>
        ${tableHtml(["Region", "Province", "Municipality", "Barangay", "Exposure Level", "Exposure Indicators", "Sector", "Grade Risk Level", "Over-all Risk Profile", "Observable Needs & Gaps", "Priority Needs", "Additional Information"], needsRows, "wide")}
        <h2>Section D: Local Response Coordination and Humanitarian Actors</h2>
        ${tableHtml(["Region", "Province", "Municipality", "Barangay", "Type Response Service", "Activity", "Actor Involved", "Details"], responseRows, "wide")}
        <h3>Narratives</h3>
        <p class="narrative">${escapeHtml(raw.sectionDNarratives || "")}</p>
        <h2>Section E: Other Details</h2>
        <h3>Team Members</h3>
        ${tableHtml(["Name", "Designation/Role", "Organization/Agency", "Email", "Mobile No."], memberRows)}
        <h3>LGU Contact Person</h3>
        ${tableHtml(["Name", "Designation", "Office", "Mobile Number", "Email Address"], lguRows)}
        <h3>Upload Photos and Other Key Documents</h3>
        <p>${escapeHtml(raw.googleDriveLink || LOCKED_DRIVE_UPLOAD_URL)}</p>
      </body>
    </html>`;
}

function keyValueTable(rows) {
  return `<table class="kv"><tbody>${rows.map((row) => `<tr><td>${escapeHtml(row[0])}</td><td class="${row[0] === "Description of Emergency" ? "narrative" : ""}">${escapeHtml(row[1])}</td></tr>`).join("")}</tbody></table>`;
}

function tableHtml(headers, rows, className) {
  const body = rows.length
    ? rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${headers.length}">No data added.</td></tr>`;
  return `<table class="${className || ""}"><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table>`;
}

function parseJson(value, fallback) {
  try {
    const parsed = JSON.parse(value || "");
    return parsed || fallback;
  } catch (error) {
    return fallback;
  }
}

function reportRecipients(raw) {
  const recipients = [NOTIFY_EMAIL];
  const members = parseJson(raw.sectionEJson, []);
  const lguContacts = parseJson(raw.sectionELguContactsJson, []);
  members.concat(lguContacts).forEach((item) => {
    const email = String((item && item.email) || "").trim();
    if (email && recipients.indexOf(email) === -1) recipients.push(email);
  });
  return recipients;
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
