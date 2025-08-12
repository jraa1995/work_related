// ===== FILE MANAGEMENT
function initializeFolderStructure() {
  const rootFolderName = "TAR_RIP_INVOICE_Processing";
  const subFolders = ["incoming", "processed", "validated", "rejected"];

  // create root
  let rootFolder;
  const folders = DriveApp.getFoldersByName(rootFolderName);
  if (folders.hasNext()) {
    rootFolder = folders.next();
  } else {
    rootFolder = DriveApp.createFolder(rootFolderName);
  }

  // create subs
  subFolders.forEach((folderName) => {
    const existingFolders = rootFolder.getFoldersByName(folderName);
    if (!existingFolders.hasNext()) {
      rootFolder.createFolder(folderName);
    }
  });

  return rootFolder.getId();
}

// ===== SAVE FILE

function saveToFolder(fileBlob, fileName, folderType) {
  const rootFolderId =
    getProperty("ROOT_FOLDER_ID") || initializeFolderStructure();
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const targetFolder = rootFolder.getFoldersByName(folderType).next();

  // timestamp to avoid duplicates
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );
  const uniqueFileName = `${timestamp}_${fileName}`;

  const file = targetFolder.createFile(fileBlob.setName(uniqueFileName));

  // metadata
  file.setDescription(
    `Uploaded: ${new Date()}, Original: ${fileName}, Type: ${folderType}`
  );

  return file.getId();
}

// ===== MOVING FILES BASED ON RISK

function moveToProcessedFolder(fileId, riskScore) {
  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();

  //  target folder based on risk score
  let targetFolderName;
  if (riskScore > 85) {
    targetFolderName = "validated";
  } else if (riskScore > 70) {
    targetFolderName = "processed";
  } else {
    targetFolderName = "rejected";
  }

  // folders
  const rootFolderId = getProperty("ROOT_FOLDER_ID");
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const targetFolder = rootFolder.getFoldersByName(targetFolderName).next();
  const currentFolder = file.getParents().next();

  // file
  targetFolder.addFile(file);
  currentFolder.removeFile(file);

  console.log(`Moved ${fileName} to ${targetFolderName} folder`);
}

// ===== TEXT EXTRACTION

function extractTextContent(fileId, fileName) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getBlob().getContentType();

    console.log(`Extracting content from ${fileName}, MIME type: ${mimeType}`);

    switch (mimeType) {
      case "application/pdf":
        return extractPdfText(file);
      case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
      case "application/msword":
        return extractDocxText(file);
      case "text/plain":
        return file.getBlob().getDataAsString();
      default:
        // try: convert to Google Doc for text extraction
        return extractViaGoogleDocs(file);
    }
  } catch (error) {
    console.error(`Error extracting text from ${fileName}:`, error);
    return `Error extracting text: ${error.message}`;
  }
}

// ===== PDF TEXT EXTRACTION

function extractPdfText(file) {
  // GScript has limited PDF text extraction
  // do we have Google Cloud Document AI
  try {
    const blob = file.getBlob();
    const resource = {
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS,
    };

    // convert PDF to Google Doc temporarily
    const tempDoc = Drive.Files.insert(resource, blob, { convert: true });
    const docContent = DocumentApp.openById(tempDoc.id).getBody().getText();

    // clean up temporary doc
    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    return docContent;
  } catch (error) {
    console.warn("PDF text extraction failed, returning placeholder");
    return `PDF content extraction failed: ${error.message}`;
  }
}

// ===== WORD DOC EXTRACTION
function extractDocxText(file) {
  try {
    const blob = file.getBlob();
    const resource = {
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS,
    };

    // Convert to Google Doc
    const tempDoc = Drive.Files.insert(resource, blob, { convert: true });
    const docContent = DocumentApp.openById(tempDoc.id).getBody().getText();

    // Clean up
    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    return docContent;
  } catch (error) {
    console.error("DOCX extraction failed:", error);
    return `Document content extraction failed: ${error.message}`;
  }
}

// ===== GOOG EXTRACTION TEST
function extractViaGoogleDocs(file) {
  try {
    const blob = file.getBlob();
    const resource = {
      title: `temp_${file.getName()}`,
      mimeType: MimeType.GOOGLE_DOCS,
    };

    const tempDoc = Drive.Files.insert(resource, blob, { convert: true });
    const content = DocumentApp.openById(tempDoc.id).getBody().getText();

    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    return content;
  } catch (error) {
    return `Unable to extract text: ${error.message}`;
  }
}
