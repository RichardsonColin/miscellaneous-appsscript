/*

------------------------------------------------------------------------------------------------------------------------------------------
DRIVE UTILS
------------------------------------------------------------------------------------------------------------------------------------------

*/
/**
 * Searches recursively to find a file by name in a given folder by its ID.
 *
 * @param {string} sourceFolderId   id of the folder to search
 * @param {string} fileName         name of the file to locate
  * @param {string} partialMatch     allows option to search for partial match

 * @return {object} the located file as a DriveApp initialization
 */

function findFileByName(sourceFolderId, fileName, partialMatch = false) {
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var found = false;
  var target;

  function search(sourceFolder, fileName) {
    var folders = sourceFolder.getFolders();
    var files = sourceFolder.getFiles();

    while (files.hasNext()) {
      if (found) break;
      var file = files.next();

      if (file.getName() === fileName) {
        target = file;
        found = true;
      } else if (partialMatch && file.getName().indexOf(fileName) !== -1) {
        target = file;
        found = true;
      }
    }

    while (folders.hasNext()) {
      if (found) break;
      var subFolder = folders.next();
      var folderName = subFolder.getName();
      search(subFolder, fileName);
    }

    if (found) return target;
  }

  search(sourceFolder, fileName);
  if (found) return target;
}

/**
 * Searches recursively to find a folder by name in a given folder by its ID.
 *
 * @param {string} sourceFolderId   id of the folder to search
 * @param {string} folderName       name of the folder to locate
 * @return {object} the located folder as a DriveApp initialization
 */

function findFolderByName(sourceFolderId, folderName) {
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var found = false;
  var target;

  function search(sourceFolder, folderName) {
    var folders = sourceFolder.getFolders();

    while (folders.hasNext()) {
      if (found) break;
      var subFolder = folders.next();

      if (subFolder.getName() === folderName) {
        target = subFolder;
        found = true;
      }

      search(subFolder, folderName);
    }

    if (found) return target;
  }

  search(sourceFolder, folderName);
  if (found) return target;
}

/**
 * Recursively copies a folder and all of its contents into a given target folder
 *
 * @param {object} folderToCopy the folder to copy
 * @param {object} targetFolder the folder to copy the contents into
 */

function copyFolder(folderToCopy, targetFolder) {
  var folders = folderToCopy.getFolders();
  var files = folderToCopy.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    file.makeCopy(file.getName(), targetFolder);
  }

  while (folders.hasNext()) {
    var subFolder = folders.next();
    var folderName = subFolder.getName();
    var newFolder = targetFolder.createFolder(folderName);
    copyFolder(subFolder, newFolder);
  }
}

/**
 * These scripts move a specific folder into a given folder, and removes the folder
 * from all other folders that previously contained it. For more information see:
 * - docs: https://developers.google.com/apps-script/drive/file
 * - source: https://stackoverflow.com/questions/18393932/implement-a-folder-move-function-in-google-dirve
 *
 * @param {string} sourceFolderId id of the drive folder to move
 * @param {string} targetFolderId id of the drive folder that receives the moved folder
 */

function moveFolderToFolderById(sourceFolderId, targetFolderId) {
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var currentFolders = sourceFolder.getParents();
  while (currentFolders.hasNext()) {
    var currentFolder = currentFolders.next();
    currentFolder.removeFolder(sourceFolder);
  }
  targetFolder.addFolder(sourceFolder);
}

/**
 * These scripts move a specific folder into a given folder, and removes the folder
 * from all other folders that previously contained it. For more information see:
 * - docs: https://developers.google.com/apps-script/drive/file
 * - source: https://stackoverflow.com/questions/18393932/implement-a-folder-move-function-in-google-dirve
 *
 * @param {string} sourceFolderName name of the drive folder to move
 * @param {string} destinationFolderName name of the drive folder that receives the moved folder
 * @param {boolean} [isUnique] sets whether the sourceFolderName must be unique
 */

function moveFolderToFolderByName(
  sourceFolderName,
  destinationFolderName,
  isUnique = false
) {
  var matchedFolders = DriveApp.getFoldersByName(sourceFolderName);

  if (matchedFolders.hasNext()) {
    var sourceFolder = matchedFolders.next();
    if (isUnique && matchedFolders.hasNext())
      throw new Error('Source Folder Name' + sourceFolderName + 'not unique');
  }

  matchedFolders = DriveApp.getFoldersByName(destinationFolderName);
  if (matchedFolders.hasNext()) {
    var targetFolder = matchedFolders.next();
    var targetFolderContents = targetFolder.getFolders();

    if (matchedFolders.hasNext())
      throw new Error(
        'Destination Folder Name: ' + destinationFolderName + ' not unique'
      );

    // check for existence of source folder name in destination folder before moving
    while (targetFolderContents.hasNext()) {
      if (targetFolderContents.next().getName() === sourceFolderName)
        throw new Error(
          'Source Folder Name: ' +
            sourceFolderName +
            ' not unique in destination folder'
        );
    }
  }

  var currentFolders = sourceFolder.getParents();
  while (currentFolders.hasNext()) {
    var currentFolder = currentFolders.next();
    currentFolder.removeFolder(sourceFolder);
  }
  targetFolder.addFolder(sourceFolder);
}

/**
 * Get a file's ID within a given folder
 *
 * @param {string} folderId Drive ID of the folder to search
 * @param {object} fileName name of the file to get
 * @return {string} ID of the searched file
 */

function getFileId(folderId, fileName) {
  var file = findFileByName(folderId, fileName);
  return file.getId();
}

/**
 * Get a folder's ID within a given folder
 *
 * @param {string} folderId Drive ID of the folder to search
 * @param {object} folderName name of the folder to get
 * @return {string} ID of the searched folder
 */

function getFolderId(folderId, folderName) {
  var folder = findFolderByName(folderId, folderName);
  return folder.getId();
}

/**
 * Finds and returns folder or creates and returns with option to move to parent folder
 *
 * @param {object} parentFolder Google Drive File of a source folder
 * @param {string} folderName name of the folder to find
 * @return {object} found or created Google Drive File
 */

function findOrCreateFolder(parentFolderId, folderName, moveTo = false) {
  const locatedFolder = findFolderByName(parentFolderId, folderName);
  const folder =
    locatedFolder !== undefined
      ? locatedFolder
      : DriveApp.createFolder(folderName);

  if (moveTo) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    folder.moveTo(parentFolder);
  }

  return folder;
}

/*

------------------------------------------------------------------------------------------------------------------------------------------
SHEET UTILS
------------------------------------------------------------------------------------------------------------------------------------------

*/

/**
 * Opens a given SpreadSheet and sets as active
 *
 * @param {string} ssId Id of the spreadsheet to prepare
 * @param {string} sheetName name of the sheet within the spreasheet
 * @return {object} SpreadSheet class
 */

function prepareSheet(ssId, sheetName) {
  if (!DriveApp.getFileById(ssId)) {
    // TODO: respond with sheet doesn't exist
    throw new Error("Yo sheeeet don't exist.");
  }
  const ss = SpreadsheetApp.openById(ssId);
  let sheet = ss.getSheetByName(sheetName);

  // create sheet if doesn't exist
  if (sheet === null) sheet = ss.insertSheet(sheetName);

  return sheet;
}

/**
 * Converts all sheets within a spreadsheet to a single pdf
 *
 * @param {string} ssId Id of the spreadsheet to prepare
 * @param {Array} sheetNames names of sheets wanted for pdf conversion
 * @return {object} File class
 */

function convertSheetsToPdf(ssId, sheetNames) {
  const ss = SpreadsheetApp.openById(ssId);
  //SpreadsheetApp.setActiveSpreadsheet(ss);
  const sheets = ss.getSheets();

  // hide sheets not for conversion
  for (let i = 0; i < sheets.length; i++) {
    if (!sheetNames.includes(sheets[i].getName())) {
      sheets[i].hideSheet();
    }
  }

  const pdfs = DriveApp.createFile(ss.getBlob());

  // show hidden sheets
  for (let i = 0; i < sheets.length; i++) {
    if (!sheetNames.includes(sheets[i].getName())) {
      sheets[i].showSheet();
    }
  }

  return pdfs;
}

/**
 * Converts a single within a spreadsheet to a pdf
 *
 * @param {string} ssId Id of the spreadsheet to prepare
 * @param {string} sheetName name of sheet wanted for pdf conversion
 * @return {object} File class
 */

function convertSheetToPdf(ssId, sheetName) {
  const ss = SpreadsheetApp.openById(ssId);
  //SpreadsheetApp.setActiveSpreadsheet(ss);
  const sheets = ss.getSheets();

  // hide sheets not for conversion
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== sheetName) {
      sheets[i].hideSheet();
    }
  }

  const pdf = DriveApp.createFile(ss.getBlob());

  // show hidden sheets
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== sheetName) {
      sheets[i].showSheet();
    }
  }

  return pdf;
}

/**
 * Write n rows to a sheet
 *
 * @param {object} sheet SpreadSheet Sheet class
 * @param {Array} pos list of position values corresponding to start of sheet write position
 * @param {number} pos[0] starting row position
 * @param {number} pos[1] starting column position
 * @param {Array} records list of data to write to rows
 * @return {object} SpreadSheet Range class
 */

function writeRows(sheet, pos, records) {
  if (!records.length) records = [[]];
  const rowLength = records[0].length;
  let valuesRange;

  if (rowLength) {
    valuesRange = sheet.getRange(pos[0], pos[1], records.length, rowLength);

    // write rows
    valuesRange.setValues(records);
  } else {
    // get range to starting position
    valuesRange = sheet.getRange(pos[0], pos[1]).activate();
  }

  return valuesRange;
}

/**
 * Hides all rows in a range that have either all 0s or emtpy strings
 *
 * @param {object} sheet SpreadSheet Sheet class
 * @param {Array} rangeValues list of A1 notation range values
 * @return {object} SpreadSheet Range class
 */

function hideEmptyRows(sheet, rangeValues) {
  const ranges = sheet.getRangeList(rangeValues).getRanges();

  for (let range of ranges) {
    let rows = range.getValues();

    for (let i = 0; i < rows.length; i++) {
      let hasData = rows[i].filter(function (cell) {
        return cell.length || cell > 0;
      });
      if (!hasData.length) {
        let emptyRowIndex = range.getRow() + i;
        sheet.hideRows(emptyRowIndex);
      }
    }
  }

  return ranges;
}

/**
 * Organize a list of data in the same order as they appear as column headings - assumes data object keys match sheet headings
 *
 * @param {object} sheet SpreadSheet Sheet class
 * @param {Array} headingPos list of position values corresponding to start of sheet headings
 * @param {Array} data list of data objects
 * @param {number} headingPos[0] starting row position
 * @param {number} headingPos[1] starting column position
 * @param {boolean} [useMax] determines whether to get max sheet columns or use data for rowLength
 * @return {Array} list of values in the order of the heading column positions
 */

function orderValuesToHeadings(sheet, headingPos, data, useMax = false) {
  if (!data.length) data = [{}];
  const rowLength = useMax
    ? sheet.getMaxColumns()
    : Math.floor(Object.keys(data[0]).length);
  let orderedData = [[]];

  if (rowLength) {
    // assumes header is only one column in height
    const headings = sheet
      .getRange(headingPos[0], headingPos[1], 1, rowLength)
      .getValues()[0];
    orderedData = data.map((record) =>
      headings.map((heading) => record[heading] || '')
    );
  }

  return orderedData;
}

/*

------------------------------------------------------------------------------------------------------------------------------------------
GENERIC HELPERS UTILS
------------------------------------------------------------------------------------------------------------------------------------------

*/

/**
 * Removes all duplicate values from an array
 *
 * @param {Array} array list of values
 * @return {Array} list of all unique values
 */

function unique(array) {
  if (Array.isArray(array) && array.length) {
    return Array.from(new Set(array));
  } else if (Array.isArray(array)) {
    return array;
  }
}

/**
 * General purpose http request fn
 *
 * @param {string} url the url to request
 * @param {object} options any added options for the request. e.g. headers, method, etc.
 * @param {string} type sets how to parse the response
 * @return {object} response to the request
 */

function fetch(url, options = {}, type = 'text') {
  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() === 200) {
    const content = response.getContentText();

    if (content !== undefined && content.length) {
      let data;

      if (type === 'json') data = JSON.parse(content);
      if (type === 'text') data = content;

      return data;
    }
  } else {
    Logger.log(url + ' Status: ' + response.getResponseCode());
    return response;
  }
}

/**
 * Fetches list of PDF data URLs
 *
 * @param {Array} data list of PDF URLs to fetch
 * @param {Array} rateLimit array containing a limit on number of requests and the time to wait
 * @param {number} rateLimit[0] number of requests allowable at once
 * @param {number} rateLimit[1] time to wait in milliseconds
 * @return {Array} list of all fetched resources as blobs
 */
//TODO: refactor to make more dynamic
function fetchPDFs(data, rateLimit = []) {
  const resources = [];
  let count = 0;

  for (let item of data) {
    if (rateLimit.length && rateLimit[0] === count) {
      count = 0;
      //await sleep(rateLimit[1]);
      Utilities.sleep(rateLimit[1]);
    }

    if (item.length) {
      Logger.log('fetch: ', item);
      let res = UrlFetchApp.fetch(item);
      let blob = res.getAs('application/pdf');
      resources.push(res);
    }

    if (rateLimit.length) count++;
  }

  return resources;
}

/**
 * Fetches list of PDF data URLs
 *
 * @param {object} params key/values of params to convert to query string
 * @return {string} string of URL query params
 */

function buildQueryParams(params) {
  return Object.keys(params).reduce(function (p, e, i) {
    return (
      p +
      (i == 0 ? '?' : '&') +
      (Array.isArray(params[e])
        ? params[e].reduce(function (str, f, j) {
            return (
              str +
              e +
              '=' +
              encodeURIComponent(f) +
              (j != params[e].length - 1 ? '&' : '')
            );
          }, '')
        : e + '=' + encodeURIComponent(params[e]))
    );
  }, '');
}

/**
 * Converts object depth to one level
 *
 * @param {object} obj multi-depth leveled object to convert
 * @return {object} converted one depth level object
 */

function flattenObject(obj) {
  const finalObject = {};

  for (let i in obj) {
    if (!obj.hasOwnProperty(i)) continue;

    if (typeof obj[i] == 'object' && obj[i] !== null) {
      const flatObject = flattenObject(obj[i]);
      for (let x in flatObject) {
        if (!flatObject.hasOwnProperty(x)) continue;
        finalObject[i + '_' + x] = flatObject[x];
      }
    } else {
      finalObject[i] = obj[i];
    }
  }
  return finalObject;
}

/**
 * Converts object to multi-line csv string
 *
 * @param {object} data one depth leveled object
 * @return {string} csv string
 */

function objectToCsv(data) {
  const replacer = (key, value) => (value === null ? '' : value);
  const headings = [];
  const totalHeadings = data.map((row) => Object.keys(row)).flat();

  for (let header of totalHeadings) {
    if (!headings.includes(header)) headings.push(header);
  }

  headings.sort();

  const csv = [
    headings.join(','),
    ...data.map((row) =>
      headings
        .map((fieldName) =>
          fieldName in row
            ? JSON.stringify(row[fieldName], replacer)
                .replace(/\r?\n|\r/g, ' ')
                .replace(/\\"/g, '""')
            : ''
        )
        .join(',')
        .trim()
    ),
  ]
    .join('\r\n')
    .trim();

  return csv;
}
