function getFolderByIdStrict_(folderId) {
  if (!folderId) {
    throw new Error('Folder ID is required.');
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    throw new Error('Unable to access folder with ID ' + folderId + ': ' + err.message);
  }
}

function getOrCreateNamedSubfolder_(parentFolder, name) {
  if (!name) {
    return parentFolder;
  }
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}

function getNextVersionedName_(parentFolder, baseName, type) {
  let maxSuffix = 0;
  const suffixPattern = new RegExp('^' + escapeForRegExp_(baseName) + '(?:' + VERSION_SUFFIX_SEPARATOR + '(\\d+))?$');
  const iterator = type === 'folder' ? parentFolder.getFolders() : parentFolder.getFiles();
  while (iterator.hasNext()) {
    const item = iterator.next();
    const name = item.getName();
    const match = name.match(suffixPattern);
    if (match) {
      if (match[1]) {
        const suffixValue = parseInt(match[1], 10);
        if (!isNaN(suffixValue)) {
          maxSuffix = Math.max(maxSuffix, suffixValue);
        }
      } else {
        maxSuffix = Math.max(maxSuffix, 1);
      }
    }
  }
  if (maxSuffix === 0) {
    return baseName;
  }
  return baseName + VERSION_SUFFIX_SEPARATOR + (maxSuffix + 1);
}

function createVersionedSubfolder(parentFolder, baseName) {
  const name = getNextVersionedName_(parentFolder, baseName, 'folder');
  const folder = parentFolder.createFolder(name);
  return { folder: folder, name: name };
}

function createVersionedSpreadsheet(baseName, parentFolder) {
  const name = getNextVersionedName_(parentFolder, baseName, 'file');
  const spreadsheet = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(spreadsheet.getId());
  parentFolder.addFile(file);
  const activeParent = file.getParents();
  while (activeParent.hasNext()) {
    const parent = activeParent.next();
    if (parent.getId() !== parentFolder.getId()) {
      parent.removeFile(file);
    }
  }
  return { spreadsheet: spreadsheet, name: name };
}

function ensureDestinationFolders() {
  const folderConfig = getFolderConfig();
  const outputFolder = getFolderByIdStrict_(folderConfig.outputFolderId);
  const reportsFolder = getOrCreateNamedSubfolder_(outputFolder, folderConfig.reportsFolderName);
  const exportsFolder = getOrCreateNamedSubfolder_(outputFolder, folderConfig.exportsFolderName);
  return {
    outputFolder: outputFolder,
    reportsFolder: reportsFolder,
    exportsFolder: exportsFolder
  };
}

function convertExcelFileToSheet(file) {
  const resource = {
    name: file.getName(),
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  const copied = Drive.Files.copy(resource, file.getId(), { supportsAllDrives: true });
  return SpreadsheetApp.openById(copied.id);
}

function deleteSpreadsheetById(spreadsheetId) {
  try {
    DriveApp.getFileById(spreadsheetId).setTrashed(true);
  } catch (err) {
    // ignore
  }
}

function escapeForRegExp_(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

