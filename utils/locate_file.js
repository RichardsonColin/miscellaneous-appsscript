function locateFileId(request) {
  var { folderId, folderName, fileName } = request;
  var subFolderId = Utils.getFolderId(folderId, folderName);
  var fileId = Utils.getFileId(subFolderId, fileName);
  return JSON.stringify(fileId);
}
