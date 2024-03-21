function sendDataToFirestore(parent,databaseName,sheetName, jsonData) {
  const emailFirestore = readSingleProperty('emailFirestore')
  const keyFirestore = readSingleProperty('keyFirestore')
  const projectIdFirestore = readSingleProperty('projectIdFirestore')
  const firestore = FirestoreApp.getFirestore(emailFirestore, keyFirestore, projectIdFirestore);

  parent = "workingCalendar"
  databaseName = "workingCalendarDatabase"

  database = databaseName

  pageSize = -1
  sheetName = 'workingCalendar'
  const {items} = app.getItems({pageSize, database, sheetName})
  jsonData = JSON.stringify(items)

  // console.log(jsonData)

  path = parent + '/' + databaseName
  firestore.createDocument(path, jsonData);

}

function sendBatchDataToFirestore(parent,databaseName,sheetName, jsonData) {
  const emailFirestore = readSingleProperty('emailFirestore')
  const keyFirestore = readSingleProperty('keyFirestore')
  const projectIdFirestore = readSingleProperty('projectIdFirestore')
  const firestore = FirestoreApp.getFirestore(emailFirestore, keyFirestore, projectIdFirestore);

  parent = "workingCalendar"
  databaseName = "workingCalendarDatabase"

  database = databaseName

  pageSize = -1
  sheetName = 'workingCalendar'
  const {items} = app.getItems({pageSize, database, sheetName})
  jsonData = JSON.stringify(items)
  path = parent + '/' + databaseName

  const requests = []
  items.forEach(function(item) {
    const request = firestore.createDocument(path, item.items)
    requests.push(request)
  })

  requests.forEach(function(request) {
    firestore.runTransaction(function(transaction) {
      transaction.createDocument(request);
      return
    })
  })
}




