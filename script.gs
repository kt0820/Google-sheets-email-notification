/**
 * AFC Patient Document Expiration Notification System
 * 
 * Main function to send expiration notifications email.
 * Groups documents by type with correct expiration calculations
 */
function sendExpirationNotifications() {
  // ====== CONFIGURATION SECTION ======
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const email = "dummyemail@gmail.com"; // Change email recipient here

  // Column mapping (zero-based indexes)
  const columnIndexes = {
    mnr: 0,             // Column A - MNR
    fullName: 1,        // Column B - FULL_NAME  
    emailCaregiver: 2,  // Column C - EMAIL_CAREGIVER
    pcpForm: 4,         // Column D - PCP-FORM
    mds: 5,             // Column E - MDS
    physical: 6,        // Column F - PHYSICAL
    pa: 8,              // Column H - PRIOR_AUTHORIZATION
    isp: 9,             // Column I - INDIVIDUAL SERVICE PLAN
    annualHealth: 10,    // Column J - ANNUAL HEALTH
  };

  // Expiration rules with exact days from the date in column
  const expirationRules = {
    pcpForm: { days: 365, displayName: "PCP Form" },                    // 365 days from date
    mds: { days: 365, displayName: "MDS" },                           // 365 days from date
    physical: { days: 365, displayName: "Physical" },                 // 365 days from date
    pa: { exact: true, displayName: "Prior Authorization" },          // Same as date in column
    isp: { days: 182, displayName: "Individual Service Plan" },       // 182 days from date
    annualHealth: { days: 365, displayName: "Annual Health" },        // 365 days from date
  };

  // ====== DATA PROCESSING SECTION ======
  // Group documents by type and status
  const documentGroups = {};
  let totalDocs = 0;
  let totalExpired = 0;
  let totalCritical = 0;

  // Initialize document groups
  Object.keys(expirationRules).forEach(field => {
    documentGroups[field] = {
      displayName: expirationRules[field].displayName,
      expired: [],
      expiring: []
    };
  });

  // Helper function to format dates consistently
  function formatDate(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
  }

  // Helper function to add days to a date
  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  // Process each row (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const fullName = row[columnIndexes.fullName];
    const emailCaregiver = row[columnIndexes.emailCaregiver] || "";

    // Check each document type for this person
    for (const [field, rule] of Object.entries(expirationRules)) {
      const cellValue = row[columnIndexes[field]];

      // Skip invalid or missing values
      if (!cellValue || typeof cellValue !== "object" || 
          cellValue === "missing" || cellValue === "discharged") {
        continue;
      }

      const baseDate = new Date(cellValue);
      let expireDate;

      // Calculate expiration date based on rule
      if (rule.exact) {
        // For PRIOR_AUTHORIZATION: use the exact date from the column
        expireDate = new Date(baseDate);
      } else {
        // For all others: add specified days to the date
        expireDate = addDays(baseDate, rule.days);
      }

      const daysDiff = Math.ceil((expireDate - today) / (1000 * 60 * 60 * 24));

      const documentData = {
        name: fullName,
        status: daysDiff < 0 ? `Expired on ${formatDate(expireDate)}` : `Expires on ${formatDate(expireDate)} (${daysDiff} days)`,
        contact: emailCaregiver,
        originalDate: formatDate(baseDate),
        expiryDate: formatDate(expireDate),
        daysRemaining: daysDiff
      };

      // Categorize by expiration status
      if (daysDiff < 0) {
        documentGroups[field].expired.push(documentData);
        totalExpired++;
        totalDocs++;
      } else if (daysDiff <= 30) {
        documentGroups[field].expiring.push(documentData);
        totalCritical++;
        totalDocs++;
      }
    }
  }

  // ====== EMAIL GENERATION SECTION ======
  // Exit if no notifications needed
  if (totalExpired === 0 && totalCritical === 0) {
    Logger.log("No expiration notifications to send.");
    return;
  }

  // Build statistics summary table
  const statsTable = `
    <h3>Overall Statistics</h3>
    <table border='1' cellpadding='8' cellspacing='0' style='border-collapse:collapse; margin-bottom:30px; width:500px;'>
      <tr style='background-color:#F5F5F5; font-weight:bold;'>
        <th>Category</th>
        <th>Total</th>
      </tr>
      <tr style='text-align:center;'>
        <td>Total Documents</td>
        <td>${totalDocs}</td>
      </tr>
      <tr style='text-align:center;'>
        <td>Expired Documents</td>
        <td>${totalExpired}</td>
      </tr>
      <tr style='text-align:center;'>
        <td>Critical (30 days)</td>
        <td>${totalCritical}</td>
      </tr>
    </table>`;

  // Build document sections grouped by document type
  let documentSections = "";

  // Process each document type
  Object.entries(documentGroups).forEach(([field, group]) => {
    const expiredCount = group.expired.length;
    const expiringCount = group.expiring.length;
    
    // Skip if no documents of this type need attention
    if (expiredCount === 0 && expiringCount === 0) return;

    // Add document type header with expiration rule info
    const rule = expirationRules[field];
    const ruleText = rule.exact ? "Expires on date shown" : `Expires ${rule.days} days after original date`;
    
    documentSections += `
      <h3 style='color:#333; border-bottom:2px solid #ddd; padding-bottom:5px;'>
        ${group.displayName} Documents 
        <small style='color:#666; font-weight:normal;'>(${ruleText})</small>
      </h3>`;

    // Add EXPIRED table for this document type
    if (expiredCount > 0) {
      documentSections += `
        <div style='background-color:#FFEBEE; padding:15px; margin-bottom:15px; border-left:5px solid #F44336;'>
          <h4 style='color:#D32F2F; margin-top:0;'>EXPIRED (${expiredCount})</h4>
          <table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse; width:100%; background-color:#FFCCCC;'>
            <tr style='background-color:#FFAAAA; font-weight:bold;'>
              <th>Name</th>
              <th>Original Date</th>
              <th>Expiry Date</th>
              <th>Status</th>
              <th>Contact</th>
            </tr>`;
      
      group.expired.forEach(doc => {
        documentSections += `
            <tr>
              <td>${doc.name}</td>
              <td>${doc.originalDate}</td>
              <td>${doc.expiryDate}</td>
              <td>${doc.status}</td>
              <td>${doc.contact}</td>
            </tr>`;
      });
      
      documentSections += `
          </table>
        </div>`;
    }

    // Add EXPIRING SOON table for this document type
    if (expiringCount > 0) {
      documentSections += `
        <div style='background-color:#FFFBF0; padding:15px; margin-bottom:15px; border-left:5px solid #FF9800;'>
          <h4 style='color:#F57C00; margin-top:0;'>EXPIRING SOON (${expiringCount})</h4>
          <table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse; width:100%; background-color:#FFF9C4;'>
            <tr style='background-color:#FFEB3B; font-weight:bold;'>
              <th>Name</th>
              <th>Original Date</th>
              <th>Expiry Date</th>
              <th>Status</th>
              <th>Contact</th>
            </tr>`;
      
      group.expiring.forEach(doc => {
        documentSections += `
            <tr>
              <td>${doc.name}</td>
              <td>${doc.originalDate}</td>
              <td>${doc.expiryDate}</td>
              <td>${doc.status}</td>
              <td>${doc.contact}</td>
            </tr>`;
      });
      
      documentSections += `
          </table>
        </div>`;
    }

    // Add spacing between document types
    documentSections += `<div style='margin-bottom:25px;'></div>`;
  });

  // Compose complete HTML email
  const htmlBody = `
    <div style='font-family: Arial, sans-serif;'>
      <h2 style='color: #333;'>Document Expiry Summary Report (${formatDate(today)})</h2>
      <hr>
      ${statsTable}
      <h2 style='color: #333;'>Patient Documents</h2>
      <div style='background-color:#F0F8FF; padding:10px; margin-bottom:20px; border-radius:5px;'>
        <strong>Expiration Rules:</strong><br>
        • PCP Form, MDS, Physical, Annual Health: <em>365 days from original date</em><br>
        • Individual Service Plan: <em>182 days from original date</em><br>
        • Prior Authorization: <em>Expires on the exact date shown</em>
      </div>
      ${documentSections}
      <br>
      <p><small>This is an automated notification from AFC Patient Management System.</small></p>
    </div>`;

  // Send the notification email
  MailApp.sendEmail({
    to: email,
    subject: `AFC Patient Expiration Notification - ${formatDate(today)}`,
    body: "Please enable HTML to view this message properly.",
    htmlBody: htmlBody
  });

  Logger.log(`Expiration notification sent: ${totalExpired} expired, ${totalCritical} critical documents.`);
}

/**
 * ====== TRIGGER MANAGEMENT SECTION ======
 * 
 * Automatically creates a time-driven trigger to run sendExpirationNotifications
 * every Thursday at 8:00 AM Eastern Time.
 * 
 * Run this function ONCE to set up the weekly automation.
 */
function createWeeklyTrigger() {
  // Remove any existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendExpirationNotifications") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Deleted existing trigger.");
    }
  }

  // Create new weekly trigger
  ScriptApp.newTrigger("sendExpirationNotifications")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(8) // 8:00 AM
    .inTimezone("America/New_York") // Eastern Time
    .create();

  Logger.log("✅ Weekly trigger created: Every Thursday at 8:00 AM EST");
  Logger.log("The system will now automatically send expiration notifications weekly.");
}

/**
 * Utility function to remove all triggers for this script.
 * Use this if you want to stop the weekly automation.
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendExpirationNotifications") {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  }
  
  Logger.log(`🗑️ Removed ${deletedCount} trigger(s). Weekly notifications are now disabled.`);
}

/**
 * ====== MANUAL TEST SECTION ======
 * 
 * Use this function to test the email generation without waiting for the trigger.
 * This helps you see the email format and verify everything works correctly.
 */
function testEmailGeneration() {
  Logger.log("🧪 Running manual test of expiration notifications...");
  sendExpirationNotifications();
  Logger.log("✅ Test completed. Check your email and the execution log.");
}
