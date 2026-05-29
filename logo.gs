var LOGO_FILES = {
  kemenhub : 'kemenhub.png',
  hubdat   : 'hubdat.png',
  sktj     : 'sktj.png'
};

function getLogoFolderId() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('LOGO_FOLDER_ID') || '1K3NjyzSmvTzNnjSZjDlp2Y1Xz2rMH997';
}

// Server-side helper untuk template embed langsung di HTML
// Dipanggil via <?= getLogoSktj() ?> di file HTML
function getLogoSktj() {
  try {
    var folderId = getLogoFolderId();
    var folder = DriveApp.getFolderById(folderId);
    var found  = folder.getFilesByName(LOGO_FILES.sktj);
    if (!found.hasNext()) return '';
    var file = found.next();
    var blob = file.getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log('[getLogoSktj] Error: ' + e.message);
    return '';
  }
}
