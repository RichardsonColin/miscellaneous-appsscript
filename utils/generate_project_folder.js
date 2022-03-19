function start(request) {
  var { sourceFolderName, destinationFolderId, templateFolderName } = request;

  var template = DriveApp.getFoldersByName(templateFolderName);
  var source = DriveApp.createFolder(sourceFolderName);
  var destination = DriveApp.getFolderById(destinationFolderId);

  if (template.hasNext()) {
    Utils.copyFolder(template.next(), source);
  }

  Utils.moveFolderToFolderById(source.getId(), destinationFolderId);
  return {
    id: source.getId(),
    url: source.getUrl(),
  };
}
