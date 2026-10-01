// PUMA legacy compatibility wrapper.
//
// The old implementation parsed project names by taking everything before the
// first hyphen and wrote directly into "Tracker Config". That behavior could
// break legitimate names such as Bryson-Kleine and create incorrect tracker
// destinations. Keep this public function name only for backward compatibility.

var OPEN_PROJECTS_FOLDER_ID = '1acRZOrQUIzhoIav1Rw8d2GPosaDvNWx5';

function updateTrackerConfigFromFolders() {
  if (typeof buildTrackerConfig !== 'function') {
    throw new Error('buildTrackerConfig() is not available.');
  }
  return buildTrackerConfig();
}
