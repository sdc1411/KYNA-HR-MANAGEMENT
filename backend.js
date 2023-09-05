const CONFIG = {
  WebAppUrl: "https://script.google.com/macros/s/AKfycbw90tw50FCBUTqENy3cW86u_dJgxuIz_rufRVz72r32UwDI6vnCZuzrqGemkIL5gIKVVg/exec",
  INDEX: "index.html",
  NAME: "KYNA HR DASHBOARD",
  TITLE: "Welcome to HR Dashboard",
  PAGE_SIZE: 15,
  REVERSE: true,
  DATABASE : {
    USERS: '1U1US3obVZMpPquEEalr8lpHwlFm4u9KGWvhs45kURSI',
    LEAVES: '1_5H7EbRGjpOW5NkB6UfNSg_r8QnHOJiXsXAFkje0ZbA',
    WORKING_CALENDAR: '1tG6S-2wHwEBEi6IvV1a0oxyHjDKR1VKLPA6M-Y7XZG8',
    EMPLOYEE_INFORMATION: '1K4TtcaK0hrRa0yGv0xSI2jP4WI8WjhGrE-pgPdcx0Mk',
    RESIGNS: '1yiORiErlPn8C94CqbTQzhgPnuDChJhkStXp_7EOdUGA',
    EMPLOYEE_CONTRACT_MANAGEMENT: '1oekZ3nVo6qlWVnYXi2y6C9EcKFB3woPcjlbLfJ1d838',
  },
  SHEET_NAME : {
    USER: 'users',
    DOC_TYPE_MANAGEMENT: 'docTypeManagement',
    WORKING_TIME_TYPE: 'workingTimeType',
    WORKING_CALENDAR: 'workingCalendar',
    },
  STATUS: {
    APPROVED: "Approved",
    REJECTED: "Rejected",
    COMPLETED: "Completed",
  }
}

// chức năng convert ngày tháng từ Vue và từ Boostrap
function convertDate(data) {
  if (data.includes("-")) {
    const dateParts1 = data.split("-")
    const dateFormated1 = new Date(+dateParts1[0], dateParts1[1] - 1, +dateParts1[2]);
    // console.log(dateFormated1)
    return dateFormated1
    
  } else {
    const dateParts2 = data.split("/")
    const dateFormated2 = new Date(+dateParts2[2], dateParts2[1] - 1, +dateParts2[0]);
    // console.log(dateFormated2)
    return dateFormated2
  }
}


function doGet(e) {

  const appType = e.parameters.appType;

  if (appType == 'reviewform') {
    const appReviewForm = new SubmitReviewForm();

    if (e.parameter.taskId || e.parameter.responseId) {
      let template;
      if (e.parameter.taskId) {
        template = HtmlService.createTemplateFromFile(ConfigReviewForm.ApprovalFlowForms.AprrovalIndexForm);
        const { task, approver, approvers, status } = appReviewForm.getTaskById(e.parameter.taskId);
        template.task = task;
        template.status = status;
        template.approver = approver;
        template.approvers = approvers;
        template.url = `${appReviewForm.url}?appType=reviewform&taskId=${e.parameter.taskId}`;
      } else if (e.parameter.responseId) {
        template = HtmlService.createTemplateFromFile(ConfigReviewForm.ApprovalFlowForms.ApprovalProgressForm);
        const { task, approvers, status } = appReviewForm.getResponseById(e.parameter.responseId);
        template.task = task;
        template.status = status;
        template.approvers = approvers;
      }

      template.title = appReviewForm.title;
      template.pending = appReviewForm.pending;
      template.approved = appReviewForm.approved;
      template.rejected = appReviewForm.rejected;
      template.waiting = appReviewForm.waiting;

      const htmlOutput = template.evaluate();
      htmlOutput
        .setTitle(appReviewForm.title)
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      return htmlOutput;

    } 

  } else if (appType == 'resignform') {
    
      const appResignForm = new SubmitResignForm();
      if (e.parameter.taskId || e.parameter.responseId) {
        let template;
        if (e.parameter.taskId) {
          template = HtmlService.createTemplateFromFile(ConfigResignForm.ApprovalFlowForms.AprrovalIndexForm);
          const { task, approver, approvers, status } = appResignForm.getTaskById(e.parameter.taskId);
          template.task = task;
          template.status = status;
          template.approver = approver;
          template.approvers = approvers;
          template.url = `${appResignForm.url}?appType=resignform&taskId=${e.parameter.taskId}`;
          // console.log(template.task)
        } else if (e.parameter.responseId) {
          template = HtmlService.createTemplateFromFile(ConfigResignForm.ApprovalFlowForms.ApprovalProgressForm);
          const { task, approvers, status } = appResignForm.getResponseById(e.parameter.responseId);
          template.task = task;
          template.status = status;
          template.approvers = approvers;
          // console.log(template.task)
        }

        template.title = appResignForm.title;
        template.pending = appResignForm.pending;
        template.approved = appResignForm.approved;
        template.rejected = appResignForm.rejected;
        template.waiting = appResignForm.waiting;

        const htmlOutput = template.evaluate();
        htmlOutput
          .setTitle(appResignForm.title)
          .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        return htmlOutput;
      }
   
  } else {
    const template = HtmlService.createTemplateFromFile(CONFIG.INDEX);
    const htmlOutput = template.evaluate();
    htmlOutput
      .setTitle(CONFIG.NAME)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return htmlOutput;
  }
  return ContentService.createTextOutput('Invalid appType');
}





function include_(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}



class App {
  constructor() {
    this.dbUser = SpreadsheetApp.openById(CONFIG.DATABASE.USERS)
    this.dbLeaves = SpreadsheetApp.openById(CONFIG.DATABASE.LEAVES)
    this.dbWorkingCalendar = SpreadsheetApp.openById(CONFIG.DATABASE.WORKING_CALENDAR)
    this.dbEmployeeInformation = SpreadsheetApp.openById(CONFIG.DATABASE.EMPLOYEE_INFORMATION)
    this.dbResigns = SpreadsheetApp.openById(CONFIG.DATABASE.RESIGNS)
    this.reverse = CONFIG.REVERSE
    this.headerId = 'id'
  }

  // getAppInfo() {
  //   const data = {
  //     name: CONFIG.NAME,
  //     title: CONFIG.TITLE,
  //   }
  //   return data
  // }

  createKeys(headers) {
    return headers.map(header => header.toString().trim())
  }

  createItemObject(keys, values) {
    const item = {}
    keys.forEach((key, index) => item[key] = values[index])
    return item
  }

  checkFilters(keys, record, filters, partial = true) {
    const results = Object.entries(filters).map(([key, value]) => {
      const index = keys.indexOf(key)
      if (partial) return new RegExp(value, 'i').test(record[index])
      return record[index] == value
    })
    if (partial) return results.includes(true)
    return !results.includes(false)
  }

  generateId(keys, records) {
    if (records.length === 0) return 1
    const indexOfId = keys.indexOf(this.headerId)
    if (indexOfId === -1) throw new Error(`"${this.headerId}" column is missing in the table!`)
    return records[records.length - 1][indexOfId] + 1
  }

  createValues(keys, item, values = []) {
    return keys.map((key, index) => {
      if (item.hasOwnProperty(key)) {
        return item[key]
      } else {
        return values[index] || null
      }
    })
  }

  getDataLoginUser() {
    const ws = this.dbUser.getSheetByName(CONFIG.SHEET_NAME.USER)
    if (!ws) throw new Error(`${CONFIG.SHEET_NAME.USER} was not found in the database`)
    const [headers,...records] = ws.getDataRange().getValues()
    
    const keys = this.createKeys(headers)

    const loginUserEmail = Session.getActiveUser().getEmail()
    // const loginUserEmail = "mouriddiep@gmail.com"

    const emailColumnIndex = 0

    const loginUsers = records.filter(row => row[emailColumnIndex] === loginUserEmail ).map(values => this.createItemObject(keys, values))

    if (loginUsers.length === 0 || loginUserEmail === "") {
      return {
        success: false,
        message: `User với email ${loginUserEmail} không tìm thấy trong Danh mục Email nhân viên`
      }
    } else {
      loginUsers[0].login_email = loginUserEmail
      const loginUser = loginUsers[0]
      return {loginUser, success: true, message: `Đăng nhập thành công`}
    }
  }

  getDataLoginNonUser(data) {
    const ws = this.dbUser.getSheetByName(CONFIG.SHEET_NAME.USER)
    if (!ws) throw new Error(`${CONFIG.SHEET_NAME.USER} was not found in the database`)
    const [headers,...records] = ws.getDataRange().getValues()
    
    const keys = this.createKeys(headers)

    const items = JSON.parse(data)
    // const loginUserEmail = item.email
    const loginNonActiveUserEmail = items.item.email
    const loginUserEmail = loginNonActiveUserEmail

   
    const emailIndividualColumnIndex = 6

    const loginUsers = records.filter(row => row[emailIndividualColumnIndex] === loginNonActiveUserEmail ).map(values => this.createItemObject(keys, values))

    if (loginUsers.length === 0 || loginUserEmail === "") {
      return {
        success: false,
        message: `User với email ${loginUserEmail} không tìm thấy trong danh mục Email cá nhân của nhân viên`
      }
    } else {
      loginUsers[0].login_email = loginUserEmail
      const loginUser = loginUsers[0]
      return {loginUser, success: true, message: `Đăng nhập thành công`}
    }
  }


  getItems({ page, pageSize, database ,sheetName, filters }) {
  
    const ss = (database === 'userDatabase') ? this.dbUser : (database === 'leaveDaysDatabase') ? this.dbLeaves : (database === 'workingCalendarDatabase') ? this.dbWorkingCalendar : (database === 'resignDatabase') ? this.dbResigns : null
    const ws = ss.getSheetByName(sheetName)
    
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    if (this.reverse) records.reverse()
    if (pageSize === -1) return {
      pages: 1,
      items: records.map(record => this.createItemObject(keys, record)),
    }
    if (filters) {
      return {
        pages: 1,
        items: records
          .filter(record => this.checkFilters(keys, record, filters))
          .map(record => this.createItemObject(keys, record))
      }
    }
    // console.log(items)
    return {
      pages: Math.ceil(records.length / pageSize),
      items: records.slice((page - 1) * pageSize, (page - 1) * pageSize + pageSize).map(record => this.createItemObject(keys, record)),
    }
  }

  createItem({ database ,sheetName, item }) {
    const ss = (database === 'leaveDaysDatabase') ? this.dbLeaves : (database === 'employeeInformationDatabase') ? this.dbEmployeeInformation : null
    const ws = ss.getSheetByName(sheetName)
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    item.createdAt = new Date()
    item.id = this.generateId(keys, records)
    const values = this.createValues(keys, item)
    ws.getRange(records.length + 2, 1, 1, values.length).setValues([values])
    return {
      success: true,
      message: `Đơn ${item.id} của bạn đã được gửi thành công!`,
      data: item,
    }
  }

  updateItem({ database ,sheetName, item }) {
    const ss = (database === 'leaveDaysDatabase') ? this.dbLeaves : (database === 'employeeInformationDatabase') ? this.dbEmployeeInformation : null
    const ws = ss.getSheetByName(sheetName)
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    const filters = {}
    filters[this.headerId] = item[this.headerId]
    const index = records.findIndex(record => this.checkFilters(keys, record, filters, false))
    if (index === -1) return {
      success: false,
      message: `Item with ID "${item.id}" was not found in the database.`
    }
    // item.modifiedOn = new Date()
    delete item.createdOn
    const values = this.createValues(keys, item, records[index])
    ws.getRange(index + 2, 1, 1, values.length).setValues([values])
    return {
      success: true,
      message: `Item ${item.id} has been updated successfully!`,
      data: item,
    }
  }

  updateFindItem(findItem, item) {
    // if (item.assignedToEmails === findItem.requestedBy) throw new Error("You can't assign the request to the requestor.")
    const data = JSON.parse(findItem.dataApprovals)
    const index = data.findIndex(v => v.email === findItem.pendingOn)
    if (item.type === "Approve") {
      data[index].status = CONFIG.STATUS.APPROVED
      data[index].comments = item.comments
      data[index].timestamp = new Date()
      if (data[index + 1]) {
        item.pendingOn = data[index + 1].email
      } else {
        item.pendingOn = null
        item.status = CONFIG.STATUS.APPROVED
      }
      // item.assignedToEmails = findItem.assignedToEmails
    } else if (item.type === "Reject") {
      data[index].status = CONFIG.STATUS.REJECTED
      data[index].comments = item.comments
      data[index].timestamp = new Date()
      item.status = CONFIG.STATUS.REJECTED
      item.pendingOn = null
      // item.assignedToEmails = findItem.assignedToEmails
    } 
    // else if (item.type === "Forward") {
    //   data[index].email = item.assignedToEmails
    //   item.pendingOn = item.assignedToEmails
    //   item.assignedToEmails = findItem.assignedToEmails.replace(findItem.pendingOn, item.assignedToEmails)
    // }
    item.dataApprovals = JSON.stringify(data)
    
  }

  updateApproval({ database ,sheetName, item }) {
    const ss = (database === 'leaveDaysDatabase') ? this.dbLeaves : null
    const ws = ss.getSheetByName(sheetName)
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    const filters = {}
    filters[this.headerId] = item[this.headerId]
    const index = records.findIndex(record => this.checkFilters(keys, record, filters, false))
    if (index === -1) return {
      success: false,
      message: `Item with ID "${item.id}" was not found in the database.`
    }

    const findItem = this.createItemObject(keys, records[index])
    this.updateFindItem(findItem, item)
    // item.modifiedOn = new Date()
    const values = this.createValues(keys, item, records[index])
    ws.getRange(index + 2, 1, 1, values.length).setValues([values])
    // console.log(values)
    // console.log(keys)
    // console.log(item)
    return {
      success: true,
      message: `Item ${item.id} has been updated successfully!`,
      data: item,
    }
  }


  updateFindCompleteItem(findItem, item) {
    const data = JSON.parse(findItem.dataHandover)
    const index = data.findIndex(v => v.email === item.completeEmail)
    if (index === -1) {
    throw new Error("Email nhân viên bàn giao chưa chính xác");
    }
    if (item.type === "Complete") {
      data[index].status = CONFIG.STATUS.COMPLETED
      data[index].comments = item.comments
      data[index].timestamp = new Date()
      if (data[index + 1]) {
        if(data[index + 1].status === "Completed") {
          item.status = CONFIG.STATUS.COMPLETED
        }
      } else if (data[index - 1]) {
        if (data[index - 1].status === "Completed") {
          item.status = CONFIG.STATUS.COMPLETED
        }
      }
    }
    item.dataHandover = JSON.stringify(data)
  }

  updateCompleteHandover({ database ,sheetName, item }) {
    const ss = (database === 'resignDatabase') ? this.dbResigns : null
    const ws = ss.getSheetByName(sheetName)
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    const filters = {}
    filters[this.headerId] = item[this.headerId]
    const index = records.findIndex(record => this.checkFilters(keys, record, filters, false))
    if (index === -1) return {
      success: false,
      message: `Item with ID "${item.id}" was not found in the database.`
    }
    // console.log(item)
    const findItem = this.createItemObject(keys, records[index])
    this.updateFindCompleteItem(findItem, item)
    const values = this.createValues(keys, item, records[index])
    ws.getRange(index + 2, 1, 1, values.length).setValues([values])
    // console.log(values)
    // console.log(keys)
    // console.log(item)
    return {
      success: true,
      message: `Biên bản bàn giao đã được cập nhật thành công!`,
      data: item,
    }
  }

  deleteItem({ database, sheetName, item }) {
    const ss = (database === 'leaveDaysDatabase') ? this.dbLeaves : (database === 'employeeInformationDatabase') ? this.dbEmployeeInformation : null
    const ws = ss.getSheetByName(sheetName)
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    const filters = {}
    filters[this.headerId] = item[this.headerId]
    const index = records.findIndex(record => this.checkFilters(keys, record, filters, false))
    if (index === -1) return {
      success: false,
      message: `Item with ID "${item.id}" was not found in the database.`
    }
    ws.deleteRow(index + 2)
    return {
      success: true,
      message: `Item ${item.id} has been deleted successfully!`,
    }
  }

}



const app = new App()

// const test = () => {
  
//   page = 1
//   pageSize = 300
//   database = 'userDatabase'
//   sheetName = 'users'
//   filters = {
//     department: 'Marketing'
//   }

//   const data = app.getItems({ page, pageSize, database ,sheetName, filters })
  
//   console.log(data)
//   return data
// }


const getDataLoginUser = (params) => JSON.stringify(app.getDataLoginUser(params))
const getDataLoginNonUser = (params) => JSON.stringify(app.getDataLoginNonUser(params))

// const getAppInfo = () => JSON.stringify(app.getAppInfo())

const getItems = (params) => JSON.stringify(app.getItems(JSON.parse(params)))
const createItem = (params) => JSON.stringify(app.createItem(JSON.parse(params)))
const updateApproval = (params) => JSON.stringify(app.updateApproval(JSON.parse(params)))
const updateCompleteHandover = (params) => JSON.stringify(app.updateCompleteHandover(JSON.parse(params)))

function logOut() {
  ScriptApp.invalidateAuth();
}






// Tạo lịch làm việc 
class GenerateWorkingCalendar {
  constructor() {
    this.dbEmployeeContractManagement = SpreadsheetApp.openById(CONFIG.DATABASE.EMPLOYEE_CONTRACT_MANAGEMENT)
    this.dbWorkingCalendar = SpreadsheetApp.openById(CONFIG.DATABASE.WORKING_CALENDAR)
  }

  getDayName(dayNumber) {
    const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    return days[dayNumber];
  }

  getWeekNumbers(dayNames) {
    const daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    return dayNames.map(dayName => daysOfWeek.indexOf(dayName));
  }

  generateWorkingTimeArray(month, year, employeeCode, department) {
    const dataSheetDocTypeManagement = this.dbEmployeeContractManagement.getSheetByName(CONFIG.SHEET_NAME.DOC_TYPE_MANAGEMENT);
    const dataDocTypeManagement = dataSheetDocTypeManagement.getDataRange().getValues();

    // console.log(dataDocTypeManagement)

    // Create an object to store the latest contract data for each employee
    const appliedDataDocTypeManagement = [];

    for (let i = 0; i < dataDocTypeManagement.length; i++) { // Start from 1 to skip header row
      const items = {
        employeeCode: dataDocTypeManagement[i][3],
        employeeName: dataDocTypeManagement[i][4],
        department: dataDocTypeManagement[i][5],
        docType: dataDocTypeManagement[i][2],
        workingTimeType: dataDocTypeManagement[i][13],
        applyFrom: dataDocTypeManagement[i][15],
        applyTo: dataDocTypeManagement[i][16],
      };

      const startDate = new Date(year, month - 1, 1);
      const endDate = new Date(year, month, 0);
      // console.log(startDate)
      // console.log(endDate)

      if (items.docType !== 'OFF' && items.workingTimeType !== '' && items.applyFrom <= endDate && items.applyTo >= startDate || items.applyTo === '') {
        appliedDataDocTypeManagement.push(items);
      }
    }

    // console.log(appliedDataDocTypeManagement)

    const workingTimeTypesSheet = this.dbEmployeeContractManagement.getSheetByName(CONFIG.SHEET_NAME.WORKING_TIME_TYPE);
    const workingTimeTypesDataRange = workingTimeTypesSheet.getRange(1, 1, workingTimeTypesSheet.getLastRow(), workingTimeTypesSheet.getLastColumn());
    const workingTimeTypesData = workingTimeTypesDataRange.getValues();

    // console.log(workingTimeTypesData)

    const data = appliedDataDocTypeManagement.map(employee => {
      const matchingWorkingTimeType = workingTimeTypesData.find(type => type[1] === employee.workingTimeType);
      
      if (!matchingWorkingTimeType) {
        console.log(`No matching working time type found for employee ${employee.employeeCode}`);
        return null; // Skip this entry
      }
      
      const workingAllDay = matchingWorkingTimeType[5] ? matchingWorkingTimeType[5].split(",") : [];
      const workingOptionalDay = matchingWorkingTimeType[8] ? matchingWorkingTimeType[8].split(",") : [];
      
      return {
        employeeCode: employee.employeeCode,
        employeeName: employee.employeeName,
        department: employee.department,
        docType: employee.docType,
        applyFrom: employee.applyFrom,
        applyTo: employee.applyTo,
        workingTimeTypeName: 'Ngày làm việc',
        workingTimeType: employee.workingTimeType,
        startShiftAllDay: matchingWorkingTimeType[3],
        endShiftAllDay: matchingWorkingTimeType[4],
        startShiftOptional: matchingWorkingTimeType[6],
        endShiftOptional: matchingWorkingTimeType[7],
        workingAllDay: this.getWeekNumbers(workingAllDay),
        workingOptionalDay: this.getWeekNumbers(workingOptionalDay),
      };
    }).filter(entry => entry !== null);

    // console.log(data)
    
    const datafilter = data.filter(item =>
          (department && employeeCode) ?
          item.department === department && item.employeeCode === employeeCode :
          (department && !employeeCode) ?
          item.department === department :
          (!department && employeeCode) ?
          item.employeeCode === employeeCode :
          true 
        );

    // console.log(datafilter)

    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);

    const workingTimeArray = [];

    const daysInMonth = endDate.getDate();

    for (let day = 1; day <= daysInMonth; day++) {
      const currentDate = new Date(year, month - 1, day);

      // console.log(datafilter.length)

      for (let i = 0; i <= datafilter.length - 1; i++) {

        // console.log(datafilter[i].applyTo)
        // console.log(currentDate > datafilter[i].applyTo)
        if (currentDate >= datafilter[i].applyFrom && currentDate <= datafilter[i].applyTo || datafilter[i].applyTo === '' ) { 
          const dayOfWeek = currentDate.getDay();
          const dayName = this.getDayName(dayOfWeek);

          const workingAllDay = datafilter[i].workingAllDay || [];
          const workingOptionalDay = datafilter[i].workingOptionalDay || [];

          let startTime = "";
          let endTime = "";
          let subid = "";

          if (workingAllDay.includes && workingAllDay.includes(dayOfWeek)) { 
              startTime = datafilter[i].startShiftAllDay;
              endTime = datafilter[i].endShiftAllDay;
              subid = "A";
              if (startTime !== "") {
                const entry = {
                  createAt: new Date(),
                  id: datafilter[i].employeeCode + currentDate.getFullYear() + currentDate.getMonth() + currentDate.getDate() + subid,
                  employeeCode: datafilter[i].employeeCode,
                  employeeName: datafilter[i].employeeName,
                  docType: datafilter[i].docType,
                  workingTimeType: datafilter[i].workingTimeType,
                  workingTimeTypeName: datafilter[i].workingTimeTypeName,
                  workingMonth: String(month),
                  workingYear: String(year),
                  workingDay: dayName,
                  startShift: startTime,
                  startDate: currentDate,
                  endShift: endTime,
                  endDate: currentDate,
                  idStatus: 'Approved',
                };
                workingTimeArray.push(entry);
              }
            }

          if (workingOptionalDay.includes && workingOptionalDay.includes(dayOfWeek)) { 
              startTime = datafilter[i].startShiftOptional;
              endTime = datafilter[i].endShiftOptional;
              subid = "O";
              if (startTime !== "") {
                const entry = {
                  createAt: new Date(),
                  id: datafilter[i].employeeCode + currentDate.getFullYear() + currentDate.getMonth() + currentDate.getDate() + subid,
                  employeeCode: datafilter[i].employeeCode,
                  employeeName: datafilter[i].employeeName,
                  docType: datafilter[i].docType,
                  workingTimeType: datafilter[i].workingTimeType,
                  workingTimeTypeName: datafilter[i].workingTimeTypeName,
                  workingMonth: String(month),
                  workingYear: String(year),
                  workingDay: dayName,
                  startShift: startTime,
                  startDate: currentDate,
                  endShift: endTime,
                  endDate: currentDate,
                  idStatus: 'Approved',
                };
                workingTimeArray.push(entry);
              }
          }
        }
      }
    }
    return workingTimeArray;
  }

  createWorkingCalendar(month, year, employeeCode, department) {
    
    const destinationSS = this.dbWorkingCalendar.getSheetByName('workingCalendar')

    const generatedData = this.generateWorkingTimeArray(month, year, employeeCode, department);
    // console.log(generatedData)

    if (generatedData.length > 0) {
      for (let i = 0; i < generatedData.length; i++) {
        const entry = generatedData[i];
        
        const duplicateRow = destinationSS
          .createTextFinder(
            `${entry.id}`
          )
          .findNext();
        
        if (!duplicateRow) {
          // Map each property to the corresponding column index in the sheet
          const row = [
            entry.createAt,
            entry.id,
            entry.employeeCode,
            entry.employeeName,
            entry.docType,
            entry.workingTimeType,
            entry.workingTimeTypeName,
            entry.workingMonth,
            entry.workingYear,
            entry.workingDay,
            entry.startShift,
            entry.startDate,
            entry.endShift,
            entry.endDate,
            entry.idStatus
          ];
          
          // console.log(row)
          // Append the row to the workingCalendarSheet
          destinationSS.appendRow(row);
          console.log("Generated sucessfull.")
            } else {
              console.log(`Lịch làm việc của nhân viên ${entry.employeeCode} trong ${entry.workingMonth}/${entry.workingYear} đã được tạo`);
            }
          }
        } else {
          console.log("No generated data to send.");}
  }

  updateWorkingCalendar(month, year, employeeCode, department) {
    
    const destinationSS = this.dbWorkingCalendar.getSheetByName('workingCalendar')
    
    const datafilter = this.generateWorkingTimeArray(month, year, employeeCode, department);

    for (const entry of datafilter) {
      const duplicateRow = destinationSS
        .createTextFinder(entry.id)
        .findNext();

      if (duplicateRow) {
        console.log(`Updating working calendar entry with id ${entry.id}`);
        
        // Map each property to the corresponding column index in the sheet
        const row = [
          entry.createAt,
          entry.id,
          entry.employeeCode,
          entry.employeeName,
          entry.docType,
          entry.workingTimeType,
          entry.workingTimeTypeName,
          entry.workingMonth,
          entry.workingYear,
          entry.workingDay,
          entry.startShift,
          entry.startDate,
          entry.endShift,
          entry.endDate,
          entry.idStatus
        ];
        
        // Get the row number of the duplicate entry
        const rowIndex = duplicateRow.getRow();
        
        // Update the existing row with the new data
        destinationSS.getRange(rowIndex, 1, 1, row.length).setValues([row]);
        
        console.log(`Updated working calendar entry with id ${entry.id}`);
      }
    }
  }

  deleteWorkingCalendar(month, year, employeeCode, department) {
    
    const destinationSS = this.dbWorkingCalendar.getSheetByName('workingCalendar')
    
    const datafilter = this.generateWorkingTimeArray(month, year, employeeCode, department);

    for (const entry of datafilter) {
      const matchingRow = destinationSS
        .createTextFinder(entry.id)
        .findNext();

      if (matchingRow) {
        console.log(`Deleting working calendar entry with id ${entry.id}`);
        
        // Get the row number of the matching entry
        const rowIndex = matchingRow.getRow();
        
        // Delete the entire row
        destinationSS.deleteRow(rowIndex);
        
        console.log(`Deleted working calendar entry with id ${entry.id}`);
      }
    }
  }

}

function createWorkingCalendar(month, year, employeeCode, department) {
  const app = new GenerateWorkingCalendar()
  app.createWorkingCalendar(month, year, employeeCode, department)
}

function updateWorkingCalendar(month, year, employeeCode, department) {
  const app = new GenerateWorkingCalendar()
  app.updateWorkingCalendar(month, year, employeeCode, department)
}

function deleteWorkingCalendar(month, year, employeeCode, department) {
  const app = new GenerateWorkingCalendar()
  app.deleteWorkingCalendar(month, year, employeeCode, department)
}

