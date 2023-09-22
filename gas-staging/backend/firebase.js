
function getDataFromFirebase (database, sheetName) {
  database = 'workingCalendarDatabase'
  sheetName = 'workingCalendarRequests'
  path = database + '/' + sheetName
  const firebaseUrl = readSingleProperty('UrlFirebase')
  const secret = readSingleProperty('SecretFirebase');
  const data = FirebaseApp.getDatabaseByUrl(firebaseUrl,secret).getData(path);
  console.log(data)
  return data
}


function writeDataFromWsToFirebase(database, sheetName) {
  // database = 'workingCalendarDatabase'
  // sheetName = 'workingCalendarRequests'
  pageSize = -1
  const {items} = app.getItems({pageSize, database, sheetName})
  const jsonData = JSON.stringify(items)
  const options = {
    method: 'put',
    contentType: 'application/json',
    payload: jsonData
  }
  
  const firebaseUrl = readSingleProperty('UrlFirebase')
  const secret = readSingleProperty('SecretFirebase');
  
  const dbFirebase = firebaseUrl + database + '/' + sheetName +'.json?auth=' +secret;
  UrlFetchApp.fetch(dbFirebase, options)
}

function pushDataFromWsToFirebase(database, sheetName) {
  // database = 'workingCalendarDatabase'
  // sheetName = 'workingCalendarRequests'
  pageSize = -1
  const {items} = app.getItems({pageSize, database, sheetName})
  const jsonData = JSON.stringify(items)
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: jsonData
  }
  
  const firebaseUrl = readSingleProperty('UrlFirebase')
  const secret = readSingleProperty('SecretFirebase');
  
  const dbFirebase = firebaseUrl + database + '/' + sheetName +'.json?auth=' +secret;
  UrlFetchApp.fetch(dbFirebase, options)
}

function updateDataFromWsToFirebase(database, sheetName) {
  // database = 'workingCalendarDatabase'
  // sheetName = 'workingCalendarRequests'

  pageSize = -1
  const {items} = app.getItems({pageSize, database, sheetName})
  const jsonData = JSON.stringify(items)
  const options = {
    method: 'patch',
    contentType: 'application/json',
    payload: jsonData
  }
  
  const firebaseUrl = readSingleProperty('UrlFirebase')
  const secret = readSingleProperty('SecretFirebase');
  
  const dbFirebase = firebaseUrl + database + '/' + sheetName +'.json?auth=' +secret;

  UrlFetchApp.fetch(dbFirebase, options)
}

function deleteDataFromWsToFirebase(database, sheetName) {
  // database = 'workingCalendarDatabase'
  // sheetName = 'workingCalendarRequests'

  const options = {
    method: 'delete',
    contentType: 'application/json',
    // payload: jsonData
  }
  
  const firebaseUrl = readSingleProperty('UrlFirebase')
  const secret = readSingleProperty('SecretFirebase');
  
  const dbFirebase = firebaseUrl + database + '/' + sheetName +'.json?auth=' +secret;

  UrlFetchApp.fetch(dbFirebase, options)
}



