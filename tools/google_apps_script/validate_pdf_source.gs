/**
 * Google Apps Script: PDF Converter Validator
 *
 * Automatically checks PDFs uploaded to Google Drive to ensure they were
 * converted using Adobe Acrobat (not third-party tools like pdf2go).
 *
 * SETUP INSTRUCTIONS:
 * 1. Open Google Drive
 * 2. Create a new Google Apps Script project (Extensions > Apps Script)
 * 3. Copy this code into the script editor
 * 4. Update FOLDER_ID with your target folder's ID
 * 5. Run setup() once to authorize permissions
 * 6. Deploy as a trigger: Edit > Current project's triggers
 *    - Function: validateNewPDFs
 *    - Event source: Time-driven
 *    - Type: Minute timer
 *    - Interval: Every 5 minutes (or your preference)
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// The Google Drive folder ID to monitor (get from folder URL)
// Example URL: https://drive.google.com/drive/folders/1ABC123XYZ
// Folder ID: 1ABC123XYZ
const FOLDER_ID = 'YOUR_FOLDER_ID_HERE';

// Email address to notify when non-Adobe PDFs are detected
const NOTIFICATION_EMAIL = 'your-email@example.com';

// Approved PDF producers (case-insensitive substring matching)
const APPROVED_PRODUCERS = [
  'Adobe PDF',
  'Adobe Acrobat',
  'Adobe InDesign',
  'Adobe Illustrator',
  'Adobe Photoshop'
];

// Action to take when non-Adobe PDF is detected
const VIOLATION_ACTION = 'MOVE_TO_QUARANTINE'; // Options: 'MOVE_TO_QUARANTINE', 'NOTIFY_ONLY', 'DELETE'

// Quarantine folder name (created automatically if using MOVE_TO_QUARANTINE)
const QUARANTINE_FOLDER_NAME = '_NON_ADOBE_PDFs_QUARANTINE';

// ============================================================================
// MAIN VALIDATION FUNCTION
// ============================================================================

/**
 * Validates all PDFs in the monitored folder that were added recently.
 * Run this on a time-based trigger (every 5-10 minutes).
 */
function validateNewPDFs() {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const cutoffTime = new Date(Date.now() - 10 * 60 * 1000); // Last 10 minutes

    const files = folder.getFiles();
    const violations = [];

    while (files.hasNext()) {
      const file = files.next();

      // Only check PDF files added recently
      if (file.getMimeType() !== 'application/pdf') continue;
      if (file.getDateCreated() < cutoffTime) continue;

      // Check if this PDF was converted with Adobe
      const validation = validatePDFSource(file);

      if (!validation.isApproved) {
        violations.push({
          file: file,
          producer: validation.producer,
          creator: validation.creator
        });

        // Take action based on configuration
        handleViolation(file, validation);
      }
    }

    // Send summary email if violations were found
    if (violations.length > 0) {
      sendViolationReport(violations);
    }

    Logger.log(`Validation complete. Found ${violations.length} non-Adobe PDFs.`);

  } catch (error) {
    Logger.log(`Error in validateNewPDFs: ${error}`);
    sendErrorEmail(error);
  }
}

// ============================================================================
// PDF METADATA VALIDATION
// ============================================================================

/**
 * Validates whether a PDF was created/converted using Adobe tools.
 *
 * @param {GoogleAppsScript.Drive.File} file - The PDF file to validate
 * @returns {Object} Validation result with producer/creator info
 */
function validatePDFSource(file) {
  try {
    // Get PDF metadata using Drive API (Advanced Drive Service)
    const metadata = Drive.Files.get(file.getId(), {
      fields: 'properties,description'
    });

    // Try to extract producer/creator from PDF
    // Note: Google Drive API doesn't expose PDF internal metadata directly
    // We need to use a workaround: download and parse the PDF
    const blob = file.getBlob();
    const pdfText = extractPDFMetadata(blob);

    const producer = pdfText.producer || 'Unknown';
    const creator = pdfText.creator || 'Unknown';

    // Check if producer matches approved Adobe signatures
    const isApproved = APPROVED_PRODUCERS.some(approved =>
      producer.toLowerCase().includes(approved.toLowerCase()) ||
      creator.toLowerCase().includes(approved.toLowerCase())
    );

    return {
      isApproved: isApproved,
      producer: producer,
      creator: creator,
      fileName: file.getName()
    };

  } catch (error) {
    Logger.log(`Error validating ${file.getName()}: ${error}`);
    return {
      isApproved: false,
      producer: 'Error',
      creator: 'Error',
      fileName: file.getName(),
      error: error.toString()
    };
  }
}

/**
 * Extracts producer and creator metadata from PDF blob.
 * PDFs store metadata in their header using format: /Producer (Adobe PDF Library 15.0)
 *
 * @param {Blob} blob - PDF file blob
 * @returns {Object} Extracted metadata
 */
function extractPDFMetadata(blob) {
  try {
    // Read first 8KB of PDF (metadata is typically in header)
    const bytes = blob.getBytes().slice(0, 8192);
    const text = Utilities.newBlob(bytes).getDataAsString('ISO-8859-1');

    // Extract Producer field
    const producerMatch = text.match(/\/Producer\s*\(([^)]+)\)/);
    const producer = producerMatch ? producerMatch[1] : 'Not found';

    // Extract Creator field
    const creatorMatch = text.match(/\/Creator\s*\(([^)]+)\)/);
    const creator = creatorMatch ? creatorMatch[1] : 'Not found';

    return {
      producer: cleanMetadataString(producer),
      creator: cleanMetadataString(creator)
    };

  } catch (error) {
    Logger.log(`Error extracting PDF metadata: ${error}`);
    return {
      producer: 'Extraction failed',
      creator: 'Extraction failed'
    };
  }
}

/**
 * Cleans metadata strings (removes PDF encoding artifacts)
 */
function cleanMetadataString(str) {
  return str
    .replace(/\\[nrtf]/g, ' ')  // Remove escape sequences
    .replace(/\s+/g, ' ')        // Normalize whitespace
    .trim();
}

// ============================================================================
// VIOLATION HANDLING
// ============================================================================

/**
 * Takes action when a non-Adobe PDF is detected.
 */
function handleViolation(file, validation) {
  const fileName = file.getName();
  const producer = validation.producer;

  Logger.log(`⚠️ VIOLATION: ${fileName} (Producer: ${producer})`);

  switch (VIOLATION_ACTION) {
    case 'MOVE_TO_QUARANTINE':
      moveToQuarantine(file);
      break;

    case 'DELETE':
      Logger.log(`Deleting: ${fileName}`);
      file.setTrashed(true);
      break;

    case 'NOTIFY_ONLY':
      // Just log and send email (handled by sendViolationReport)
      break;

    default:
      Logger.log(`Unknown action: ${VIOLATION_ACTION}`);
  }

  // Add comment to file for audit trail
  try {
    Drive.Comments.create({
      content: `⚠️ AUTO-FLAGGED: Not converted with Adobe Acrobat\n` +
               `Producer: ${producer}\n` +
               `Creator: ${validation.creator}\n` +
               `Detected: ${new Date().toISOString()}`
    }, file.getId());
  } catch (error) {
    Logger.log(`Could not add comment: ${error}`);
  }
}

/**
 * Moves file to quarantine folder
 */
function moveToQuarantine(file) {
  try {
    // Get or create quarantine folder
    const parentFolder = DriveApp.getFolderById(FOLDER_ID);
    let quarantineFolder;

    const existingFolders = parentFolder.getFoldersByName(QUARANTINE_FOLDER_NAME);
    if (existingFolders.hasNext()) {
      quarantineFolder = existingFolders.next();
    } else {
      quarantineFolder = parentFolder.createFolder(QUARANTINE_FOLDER_NAME);
    }

    // Move file
    file.moveTo(quarantineFolder);
    Logger.log(`Moved ${file.getName()} to quarantine`);

  } catch (error) {
    Logger.log(`Error moving to quarantine: ${error}`);
  }
}

// ============================================================================
// NOTIFICATION EMAILS
// ============================================================================

/**
 * Sends violation report email
 */
function sendViolationReport(violations) {
  if (!NOTIFICATION_EMAIL || NOTIFICATION_EMAIL === 'your-email@example.com') {
    Logger.log('No notification email configured');
    return;
  }

  const subject = `⚠️ Non-Adobe PDFs Detected (${violations.length} files)`;

  let body = `<html><body>`;
  body += `<h2>PDF Converter Validation Report</h2>`;
  body += `<p><strong>${violations.length}</strong> PDF(s) were detected that were NOT converted using Adobe Acrobat:</p>`;
  body += `<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">`;
  body += `<tr style="background-color: #f0f0f0;">`;
  body += `<th>File Name</th><th>Producer</th><th>Creator</th><th>Action Taken</th>`;
  body += `</tr>`;

  violations.forEach(v => {
    body += `<tr>`;
    body += `<td>${v.file.getName()}</td>`;
    body += `<td>${v.producer}</td>`;
    body += `<td>${v.creator}</td>`;
    body += `<td>${VIOLATION_ACTION}</td>`;
    body += `</tr>`;
  });

  body += `</table>`;
  body += `<br><p><strong>Expected Producer:</strong> ${APPROVED_PRODUCERS.join(', ')}</p>`;
  body += `<p><strong>Folder:</strong> <a href="https://drive.google.com/drive/folders/${FOLDER_ID}">View Folder</a></p>`;
  body += `<hr><p style="color: #666; font-size: 12px;">`;
  body += `This is an automated message from Google Apps Script PDF Validator.<br>`;
  body += `Time: ${new Date().toLocaleString()}`;
  body += `</p></body></html>`;

  try {
    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      htmlBody: body
    });
    Logger.log(`Notification email sent to ${NOTIFICATION_EMAIL}`);
  } catch (error) {
    Logger.log(`Error sending email: ${error}`);
  }
}

/**
 * Sends error notification email
 */
function sendErrorEmail(error) {
  if (!NOTIFICATION_EMAIL || NOTIFICATION_EMAIL === 'your-email@example.com') {
    return;
  }

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: '❌ PDF Validator Script Error',
    body: `The PDF validator script encountered an error:\n\n${error}\n\nTime: ${new Date()}`
  });
}

// ============================================================================
// MANUAL TESTING & SETUP
// ============================================================================

/**
 * Run this once to set up and test the script
 */
function setup() {
  Logger.log('=== PDF Validator Setup ===');

  // Test folder access
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    Logger.log(`✓ Folder access OK: ${folder.getName()}`);
  } catch (error) {
    Logger.log(`❌ Cannot access folder: ${error}`);
    return;
  }

  // Enable Advanced Drive Service
  Logger.log('\n⚠️ IMPORTANT: Enable "Drive API" in Services:');
  Logger.log('1. Click "+" next to Services in left sidebar');
  Logger.log('2. Find "Drive API" and click "Add"');

  Logger.log('\n✓ Setup complete!');
  Logger.log('\nNext steps:');
  Logger.log('1. Set up a time-based trigger for validateNewPDFs()');
  Logger.log('2. Upload a test PDF and wait for the trigger to run');
}

/**
 * Manually test validation on all PDFs in folder (for debugging)
 */
function testValidation() {
  Logger.log('=== Manual Validation Test ===\n');

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType('application/pdf');

  let count = 0;
  while (files.hasNext() && count < 5) { // Test first 5 PDFs
    const file = files.next();
    Logger.log(`Testing: ${file.getName()}`);

    const validation = validatePDFSource(file);
    Logger.log(`  Producer: ${validation.producer}`);
    Logger.log(`  Creator: ${validation.creator}`);
    Logger.log(`  Approved: ${validation.isApproved ? '✓ YES' : '✗ NO'}`);
    Logger.log('');

    count++;
  }
}

/**
 * View current configuration
 */
function showConfig() {
  Logger.log('=== Current Configuration ===');
  Logger.log(`Folder ID: ${FOLDER_ID}`);
  Logger.log(`Notification Email: ${NOTIFICATION_EMAIL}`);
  Logger.log(`Approved Producers: ${APPROVED_PRODUCERS.join(', ')}`);
  Logger.log(`Violation Action: ${VIOLATION_ACTION}`);
  Logger.log(`Quarantine Folder: ${QUARANTINE_FOLDER_NAME}`);
}
