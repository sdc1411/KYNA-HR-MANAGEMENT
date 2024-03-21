 
function saveSingleProperty(key, value) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(key, value);
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


function saveMultipleProperties() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperties({
      // 'SecretFirebase': '',
      // 'UrlFirebase': ''
      // emailFirestore: '',
      // keyFirestore: '',
      // projectIdFirestore: '',
      // serverTimeSheet: '',
      // passwordTimeSheet: '',
      // usernameTimeSheet: '',
      // databaseTimeSheet: '',
    });
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


function readSingleProperty(key) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const value = scriptProperties.getProperty(key);
    return value
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


function readAllProperties() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const data = scriptProperties.getProperties();
    for (const key in data) {
      console.log('Key: %s, Value: %s', key, data[key]);
    }
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


function updateProperty(key, newValue) {
  key =  ''
  newValue = ''
      // keyFirestore: '',
      // projectIdFirestore: '',
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let units = scriptProperties.getProperty(key);
    units = newValue; // Only changes local value, not stored value.
    scriptProperties.setProperty(key, units); // Updates stored value.
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


function deleteSingleProperty(key) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty(key);
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


function deleteAllUserProperties() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}