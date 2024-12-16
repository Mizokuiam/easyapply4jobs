![YT Thumbnail](https://github.com/user-attachments/assets/21c1d253-b437-4fbc-a792-2d0a55f767ec)

لكود:
function generateAndSendEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const appProfileSheet = sheet.getSheetByName("App_Profile");
  const actionTableSheet = sheet.getSheetByName("Action_Table");
  const emailTemplateSheet = sheet.getSheetByName("Email_Template");

  const coverLetterFolderName = "Generated Cover Letters";
  const coverLetterFolder = getOrCreateFolder(coverLetterFolderName);

  const appProfileData = appProfileSheet.getDataRange().getValues();
  const actionTableData = actionTableSheet.getDataRange().getValues();
  const emailTemplate = emailTemplateSheet ? emailTemplateSheet.getRange(2, 1, 1, 2).getValues()[0] : ["Applying for {JobPosition}", ""]; // Default subject and body

  // Check for and delete generated files older than 30 days
  deleteOldGeneratedFiles(coverLetterFolder, 30);

  for (let i = 1; i < actionTableData.length; i++) {
    const [jobPosition, companyName, contactEmail, emailStatus, generatedStatus, emailDateTime, attachmentLink] = actionTableData[i];

    // Process only if the status is not generated yet
    if (generatedStatus !== "Generated") {
      const appProfile = appProfileData.find(row => row[0] === jobPosition);
      if (!appProfile) {
        console.log(`No matching job profile found for ${jobPosition}`);
        continue;
      }

      const [position, templateLink, resumeFileName] = appProfile;
      const templateId = getIdFromLink(templateLink);
      const resumeFile = getFileByName(resumeFileName);

      if (!resumeFile) {
        console.log(`Resume file not found: ${resumeFileName}`);
        continue;
      }

      try {
        console.log(`Generating cover letter for ${companyName} (${jobPosition})`);
        const coverLetterFile = generateCoverLetter(templateId, coverLetterFolder, { CompanyName: companyName, JobPosition: jobPosition, ContactEmail: contactEmail });
        const coverLetterLink = coverLetterFile.getUrl();

        // Update the sheet with the generated details
        actionTableSheet.getRange(i + 1, 7).setValue(coverLetterLink); // Add the link to the "Generated Cover Letter (Attachment)" column (Column 7)
        actionTableSheet.getRange(i + 1, 5).setValue("Generated"); // Set "Generated" status in column 5 (Generated Status)

        // Prepare email body and subject
        const subject = emailTemplate[0].replace("{JobPosition}", jobPosition);
        const body = emailTemplate[1]
          .replace("{CompanyName}", companyName)
          .replace("{JobPosition}", jobPosition)
          .replace("{ContactEmail}", contactEmail);

        // Send email
        if (emailStatus !== "Sent") {
          console.log(`Sending email to ${contactEmail} for ${jobPosition}`);
          sendEmail(contactEmail, subject, body, coverLetterFile, resumeFile);

          // Log email sent date and time in "Email Date/Time" column (column 6 in Action_Table)
          const emailDateTime = new Date();
          actionTableSheet.getRange(i + 1, 4).setValue("Sent"); // Update the "Email Status" column (Column 4)
          actionTableSheet.getRange(i + 1, 6).setValue(emailDateTime); // Update the "Email Date/Time" column (Column 6)
        }
      } catch (error) {
        console.error(`Error processing ${companyName}: ${error.message}`);
      }
    }
  }
}

function sendEmail(to, subject, body, coverLetterFile, resumeFile) {
  // Use <br> for line breaks in the HTML body
  const htmlBody = body.replace(/\n/g, "<br>");

  GmailApp.sendEmail(to, subject, body, {
    htmlBody: htmlBody,
    attachments: [coverLetterFile, resumeFile],
  });
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next(); // Return the existing folder
  }
  // Create the folder if it doesn't exist
  return DriveApp.createFolder(folderName);
}

// Additional helper functions
function getFileByName(fileName) {
  const files = DriveApp.getFilesByName(fileName);
  if (files.hasNext()) {
    return files.next();
  }
  return null;
}

function getIdFromLink(link) {
  const match = link.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function generateCoverLetter(templateId, folder, placeholders) {
  const templateFile = DriveApp.getFileById(templateId);
  const templateCopy = templateFile.makeCopy(`CoverLetter_${placeholders.JobPosition}_${placeholders.CompanyName}`, folder);
  const doc = DocumentApp.openById(templateCopy.getId());
  const body = doc.getBody();

  // Replace placeholders
  for (let key in placeholders) {
    body.replaceText(`{${key}}`, placeholders[key]);
  }
  doc.saveAndClose();

  // Export to PDF
  const pdfFile = DriveApp.getFileById(templateCopy.getId()).getAs("application/pdf");
  const pdfCopy = folder.createFile(pdfFile).setName(`CoverLetter_${placeholders.JobPosition}_${placeholders.CompanyName}.pdf`);
  
  // Clean up the template copy
  templateCopy.setTrashed(true);

  // Move the generated PDF file to the "Generated Cover Letters" folder
  const coverLetterFolder = getOrCreateFolder("Generated Cover Letters");
  coverLetterFolder.addFile(pdfCopy);
  DriveApp.getRootFolder().removeFile(pdfCopy); // Remove it from the root folder

  return pdfCopy;
}

function deleteOldGeneratedFiles(folder, days) {
  const files = folder.getFiles();
  const currentDate = new Date();

  while (files.hasNext()) {
    const file = files.next();
    const createdDate = file.getDateCreated();
    const diffInTime = currentDate - createdDate;
    const diffInDays = diffInTime / (1000 * 3600 * 24); // Convert milliseconds to days

    if (diffInDays > days) {
      console.log(`Deleting file: ${file.getName()} (Created: ${createdDate})`);
      file.setTrashed(true); // Move the file to trash
    }
  }
}

