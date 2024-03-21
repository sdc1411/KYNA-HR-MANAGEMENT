class HandleFirestore {
  constructor() {
    this.firestore = this.connectFirestore();
  }

  connectFirestore() {
    const emailFirestore = readSingleProperty('emailFirestore')
    const keyFirestore = readSingleProperty('keyFirestore')
    const projectIdFirestore = readSingleProperty('projectIdFirestore')
    const firestore = FirestoreApp.getFirestore(emailFirestore, keyFirestore, projectIdFirestore);
    return firestore
  }

  sendMultiDataToFirestore({path, datas}) {
    try {
      for ( let i = 0; i < datas.length; i++) {
        this.firestore.createDocument(path, datas[i]);
      }
    } catch (error) {
      console.error('Error creating document:', path, error);
    }
  }

  sendDataToFirestore({path, data}) {
    try {
      this.firestore.createDocument(path, data);
    } catch (error) {
      console.error('Error creating document:', path, error);
    }
  }

  getDataFromFirestore(path) {
    try {
      const documents = this.firestore.getDocuments(path);
      documents.forEach((doc) => {
        console.log(doc.name, " => ", doc.fields);
      });
    } catch (err) {
      console.error('Không tìm thấy data')
    }
  }

  queryDataFromFirestore({path, key, value}) {
    const allDocumentsNullNames = this.firestore.query(path).Where(key, value).Execute();
  }

  updateDataToFirestore({path, data}) {
    try {
      this.firestore.updateDocument(path, data);
    } catch (error) {
      console.error('Error creating document:', path, error);
    }
  }

  deleteData (path) {
    try {
      this.firestore.deleteDocument(path);
    } catch (error) {
      console.error('Error creating document:', path, error);
    }
  }

}

function testDataFromFirestore () {
  const path = "messages"
  const app = new HandleFirestore()
  app.getDataFromFirestore(path)
}


function getWorkingCalendarDataFromFirestore () {
  let collection = "workingCalendar"
  let workingMonth = 9
  let workingYear = 2023
  let departmentCode = 4
  let idMonth = (workingMonth < 10) ? '0' + workingMonth : workingMonth
  const yearMonth = workingYear + idMonth
  const path = collection + '/' + yearMonth + '/' + departmentCode
  const app = new HandleFirestore()
  app.getDataFromFirestore(path)
}

function deleteWorkingCalendarDataFromFirestore () {
  let collection = "workingCalendar"
  let workingMonth = 9
  let workingYear = 2023
  let departmentCode = 1
  let idMonth = (workingMonth < 10) ? '0' + workingMonth : workingMonth
  const yearMonth = workingYear + idMonth
  const path = collection + '/' + yearMonth + '/' + departmentCode
  const app = new HandleFirestore()
  app.deleteData(path)
}

function importWorkingCalendarDataToFirestore() {
  let collection = "workingCalendar"
  let database = "workingCalendarDatabase"
  let sheetName = 'workingCalendar'
  let workingMonth = 9
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
      const yearMonth = workingYear + idMonth
      const departmentCode = getDepartmentCode(listDepartment[i])
      const path = collection + '/' + yearMonth + '/' + departmentCode
      
      const app = new HandleFirestore()
      app.sendMultiDataToFirestore({path: path,datas: result})
    }
  }
}


function importUsersDataToFirestore() {
  let collection = "users"
  let database = "userDatabase"
  let sheetName = 'users'

  const getItems = new App()
  const { items } = getItems.getItems({ database: database,sheetName: sheetName, pageSize: -1})
  const result = items.filter(item => item.Employee_code !== '')
  // console.log(result)
  
  for (let i = 0; i < result.length; i++) {
    const path = collection + '/' + result[i].Employee_code
    // console.log(path, result[i])
    // console.log(result[i])
    const app = new HandleFirestore()
    app.sendDataToFirestore({path: path,data: result[i]})
  }
}





