class HandleFirebase {
  constructor() {
    this.firebaseUrl = readSingleProperty('UrlFirebase');
    this.secret = readSingleProperty('SecretFirebase');
  }
  
  getData (path) {
    const result = FirebaseApp.getDatabaseByUrl(this.firebaseUrl,this.secret).getData(path);
    return result
  }

  getDataWithQuery (path, query) {
    const result = FirebaseApp.getDatabaseByUrl(this.firebaseUrl,this.secret).getData(path, query);
    return result
  }

  getAllData (paths) {
    const result = FirebaseApp.getDatabaseByUrl(this.firebaseUrl,this.secret).getAllData(paths);
    return result
  }

  writeData({path, data}) {
    const options = {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(data)
    }
    
    const dbFirebase = this.firebaseUrl + path +'.json?auth=' + this.secret;
    UrlFetchApp.fetch(dbFirebase, options)
  }

  pushData ({path, data}) {
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(data)
    }
    
    const dbFirebase = this.firebaseUrl + path +'.json?auth=' + this.secret;
    UrlFetchApp.fetch(dbFirebase, options)
  }

  deleteData(path) {
    FirebaseApp.getDatabaseByUrl(this.firebaseUrl,this.secret).removeData(path);
  }

  updateData({path, data}) {
    const jsonData = JSON.stringify(data)
    // console.log(jsonData)
    const result = FirebaseApp.getDatabaseByUrl(this.firebaseUrl,this.secret).updateData(path,jsonData);
    return result
  }

}

function deleteWorkingCalendarDataFromFirebase() {
  let collection = "workingCalendar"
  let workingMonth = 10
  let workingYear = 2023
  let idMonth = (workingMonth < 10) ? '0' + workingMonth : workingMonth
  const yearMonth = workingYear + idMonth
  const path = collection

  const app = new HandleFirebase()
  app.deleteData(path)

}

function getWorkingCalendarFromFirebase(collection, month) {
  collection = "leaveTracking"
  month = '202310'
  const path = collection + '/' + month

  const app = new HandleFirebase()
  const resutl = app.getData(path)
  // console.log(Array.isArray(data))
  return JSON.stringify(result)
}

function getWorkingCalendarFromFirebaseWithQuery() {
  const collection = "userDatabase/users"
  const path = collection

  const app = new HandleFirebase()
  const resutl = app.getData(path)
  console.log(resutl)
  return JSON.stringify(resutl)
}

function getAllWorkingCalendarFromFirebase(params) {
  const months = ["202310","202309","202311"]
  const collections = ["workingCalendar","leaveTracking","timeSheet"]

  const paths = []
  for (let i = 0; i < collections.length; i++) {
    for (let j = 0; j < months.length; j++) {
      const path = collections[i] + '/' + months[j]
      paths.push(path)
    }
  }
  const app = new HandleFirebase()
  const result = app.getAllData(paths).flat(2).filter(item => item !== null)
  // console.log(data)
  // console.log(JSON.stringify(data))
  const response = {
    items: result
  };
  return JSON.stringify(response);
}


function importWorkingCalendarDataToFiresbase() {
  let collection = "leaveTracking"
  let database = "leaveDaysDatabase"
  let sheetName = 'leaveTracking'
  let workingMonth = 10
  let workingYear = 2023
  let filters = {
    workingMonth: workingMonth,
    workingYear: workingYear,
  }

  const getItems = new App()
  const { items } = getItems.getItems({ database: database,sheetName: sheetName, filters: filters})

  const listDepartment = [
    'Board',
    'Academic - Curriculum',
    'Academic - Student Care',
    'Academic - Teacher Care',
    'Accounting & Finance',
    'Business Development',
    'Course Design',
    'Cusomer Service (Video)',
    'Design',
    'Admin Logistics',
    'Marketing',
    'Math - Curriculum',
    'Math - Teacher Care',
    'Math - Student Care',
    'Operation',
    'Sales Tutoring',
    'Sales Video',
    'Technology & IT',
    'Teacher',
    'Human Resources & Adminisrtation']
  
  for (let i = 0; i < listDepartment.length; i++) {
    const result = items.filter(item => item.department === listDepartment[i] && item.workingMonth === workingMonth && item.workingYear === workingYear)
    if (result.length > 0) {
      let idMonth = (workingMonth < 10) ? '0' + workingMonth : workingMonth
      const yearMonth = String(workingYear) + String(idMonth)
      const departmentCode = getDepartmentCode(listDepartment[i])
      const path = collection + '/' + yearMonth + '/' + departmentCode
      
      const app = new HandleFirebase()
      app.writeData({path: path,data: result})
    }
  }
}

function updateWorkingCalendarDataToFiresbase() {
  let collection = "leaveTracking"
  let database = "leaveDaysDatabase"
  let sheetName = 'leaveTracking'
  let workingMonth = 10
  let workingYear = 2023
  let filters = {
    workingMonth: workingMonth,
    workingYear: workingYear,
  }

  const getItems = new App()
  const { items } = getItems.getItems({ database: database,sheetName: sheetName, filters: filters})

  const listDepartment = [
    'Board',
    'Academic - Curriculum',
    'Academic - Student Care',
    'Academic - Teacher Care',
    'Accounting & Finance',
    'Business Development',
    'Course Design',
    'Cusomer Service (Video)',
    'Design',
    'Admin Logistics',
    'Marketing',
    'Math - Curriculum',
    'Math - Teacher Care',
    'Math - Student Care',
    'Operation',
    'Sales Tutoring',
    'Sales Video',
    'Technology & IT',
    'Teacher',
    'Human Resources & Adminisrtation']
  
  for (let i = 0; i < listDepartment.length; i++) {
    const result = items.filter(item => item.department === listDepartment[i] && item.workingMonth === workingMonth && item.workingYear === workingYear)
    if (result.length > 0) {
      let idMonth = (workingMonth < 10) ? '0' + workingMonth : workingMonth
      const yearMonth = String(workingYear) + String(idMonth)
      const departmentCode = getDepartmentCode(listDepartment[i])
      const path = collection + '/' + yearMonth + '/' + departmentCode
      
      const app = new HandleFirebase()
      app.pushData({path: path,data: result})
    }
  }
}

