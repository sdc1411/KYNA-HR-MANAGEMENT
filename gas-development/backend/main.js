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

function  getDepartmentCode(departmentName) {
  const departmentCodeMapping = {
    'Board': 1,
    'Academic - Curriculum': 2,
    'Academic - Student Care': 3,
    'Academic - Teacher Care': 4,
    'Accounting & Finance': 5,
    'Business Development': 6,
    'Course Design': 7,
    'Cusomer Service (Video)': 8,
    'Design': 9,
    'Admin Logistics': 11,
    'Marketing': 12,
    'Math - Curriculum': 13,
    'Math - Teacher Care': 15,
    'Math - Student Care': 14,
    'Operation': 16,
    'Sales Tutoring': 17,
    'Sales Video': 18,
    'Technology & IT': 19,
    'Teacher': 20,
    'Human Resources & Adminisrtation': 10
  };
  return departmentCodeMapping[departmentName] || departmentName;
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
    return htmlOutput;
  }
  return ContentService.createTextOutput('Invalid appType');
}





function include_(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}



class App {
  constructor() {
    this.reverse = CONFIG.REVERSE
    this.headerId = 'id'
  }

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
    const dbUser = SpreadsheetApp.openById(CONFIG.DATABASE.USERS)
    const wsUsers = dbUser.getSheetByName(CONFIG.SHEET_NAME.USER)
    const [headers,...records] = wsUsers.getDataRange().getValues()
    
    const keys = this.createKeys(headers)

    const loginUserEmail = Session.getActiveUser().getEmail()

    const emailColumnIndex = 0

    const loginUsers = records.filter(row => row[emailColumnIndex] === loginUserEmail ).map(values => this.createItemObject(keys, values))

    if (loginUsers.length === 0 || loginUserEmail === "") {
      return {
        success: false,
        message: `User với email ${loginUserEmail} không tìm thấy trong Danh mục Email nhân viên`
      }
    } else {
      loginUsers[0].login_email = loginUserEmail
      const wsLeaveRequests = dbUser.getSheetByName(CONFIG.SHEET_NAME.LEAVE_REQUESTS)
      const [leaveRequestHeaders,...leaveRequestRecords] = wsLeaveRequests.getDataRange().getValues()
      const leaveRequestKeys = this.createKeys(leaveRequestHeaders)

      const wsLeaveTypes = dbUser.getSheetByName(CONFIG.SHEET_NAME.LEAVE_TYPES)
      const [leaveTypeHeaders,...leaveTypeRecords] = wsLeaveTypes.getDataRange().getValues()
      const leaveTypeKeys = this.createKeys(leaveTypeHeaders)

      return {
        loginUser: loginUsers[0], 
        dataUsers: {pages: 1, items: records.map(record => this.createItemObject(keys, record))},
        dataLeaveRequests: {pages: 1, items: leaveRequestRecords.map(record => this.createItemObject(leaveRequestKeys, record))},
        dataLeaveTypes: {pages: 1, items: leaveTypeRecords.map(record => this.createItemObject(leaveTypeKeys, record))},
        app: {name: CONFIG.NAME, title: CONFIG.TITLE},
        success: true,
        message: `Đăng nhập thành công`}
    }
  }

  getDataLoginNonUser(data) {
    const dbUser = SpreadsheetApp.openById(CONFIG.DATABASE.USERS)
    const ws = dbUser.getSheetByName(CONFIG.SHEET_NAME.USER)
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

  sendNotificationApproval(sheetName, id, employeeName, department, pendingOn) {
    if (sheetName === 'leaveRequests') {
      const title = 'Đơn xin nghỉ phép, WFH'
      const template = HtmlService.createTemplateFromFile(CONFIG.NOTIFICATION_APPROVAL.Template)
      template.id = id
      template.employeeName = employeeName
      template.department = department
      template.title = title
      template.url = CONFIG.WebAppUrl
      
      const subject = `Approval Pending - ${title} - ${employeeName} - ${department}`

      const options = {
        htmlBody: template.evaluate().getContent()
      }
      GmailApp.sendEmail(pendingOn, subject, "", options);
    } else {return}
  }

  getItems({ page, pageSize, database, sheetName, filters }) {
    const id = (database === 'userDatabase') ? CONFIG.DATABASE.USERS : 
               (database === 'leaveDaysDatabase') ? CONFIG.DATABASE.LEAVES : 
               (database === 'workingCalendarDatabase') ? CONFIG.DATABASE.WORKING_CALENDAR : 
               (database === 'resignDatabase') ? CONFIG.DATABASE.RESIGNS : 
               (database === 'employeeContractDatabase') ? CONFIG.DATABASE.EMPLOYEE_CONTRACT_MANAGEMENT : 
               (database === 'timeSheetDatabase') ? CONFIG.DATABASE.TIME_SHEET : null;

    return this.getSheetItems(id, sheetName, pageSize, page, filters);
  }

  getSheetItems(id, sheetName, pageSize, page, filters) {
    const ss = SpreadsheetApp.openById(id)
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
    const id = (database === 'leaveDaysDatabase') ? CONFIG.DATABASE.LEAVES : (database === 'employeeInformationDatabase') ? CONFIG.DATABASE.EMPLOYEE_INFORMATION : (database === 'employeeContractDatabase') ? CONFIG.DATABASE.EMPLOYEE_CONTRACT_MANAGEMENT : (database === 'workingCalendarDatabase') ? CONFIG.DATABASE.WORKING_CALENDAR : (database === 'timeSheetDatabase') ? CONFIG.DATABASE.TIME_SHEET : null
    const ss = SpreadsheetApp.openById(id)
    const ws = ss.getSheetByName(sheetName)
    if (!ws) throw new Error(`${sheetName} was not found in the database.`)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    item.createdAt = new Date()
    item.id = this.generateId(keys, records)
    const values = this.createValues(keys, item)
    ws.getRange(records.length + 2, 1, 1, values.length).setValues([values])
    if (database === 'leaveDaysDatabase') {
      if (item.pendingOn) {
        this.sendNotificationApproval(sheetName,item.id,item.employeeName,item.department,item.pendingOn)
      }
      if (item.status === 'Approved') {
        const app = new PushLeaveTracking()
        app.pushLeaveTracking(item.id)
      }
    }
    return {
      success: true,
      message: `Đơn ${item.id} của bạn đã được gửi thành công!`,
      data: item,
    }
  }

  updateItem({ database ,sheetName, item }) {
    const id = (database === 'leaveDaysDatabase') ? CONFIG.DATABASE.LEAVES : (database === 'employeeInformationDatabase') ? CONFIG.DATABASE.EMPLOYEE_INFORMATION : (database === 'employeeContractDatabase') ? CONFIG.DATABASE.EMPLOYEE_CONTRACT_MANAGEMENT : null
    const ss = SpreadsheetApp.openById(id)
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
    if (item.assignedToEmails === findItem.requestedBy) throw new Error("You can't assign the request to the requestor.")
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
      item.assignedToEmails = findItem.assignedToEmails
    } 
    else if (item.type === "Forward") {
      data[index].email = item.assignedToEmails
      item.pendingOn = item.assignedToEmails
      item.assignedToEmails = findItem.assignedToEmails.replace(findItem.pendingOn, item.assignedToEmails)
    }
    item.dataApprovals = JSON.stringify(data)
    
  }

  updateApproval({ database ,sheetName, item }) {
    const id = (database === 'leaveDaysDatabase') ? CONFIG.DATABASE.LEAVES : null
    const ss = SpreadsheetApp.openById(id)
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
    if (database === 'leaveDaysDatabase') {
      if (item.pendingOn) {
        this.sendNotificationApproval(sheetName,item.id,item.employeeName,item.department,item.pendingOn)
      }
      if (item.status === 'Approved') {
        const app = new PushLeaveTracking()
        app.pushLeaveTracking(item.id)
      }
    }
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
    const id = (database === 'resignDatabase') ? CONFIG.DATABASE.RESIGNS : null
    const ss = SpreadsheetApp.openById(id)
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
    const id = (database === 'leaveDaysDatabase') ? CONFIG.DATABASE.LEAVES : (database === 'employeeInformationDatabase') ? CONFIG.DATABASE.EMPLOYEE_INFORMATION : null
    const ss = SpreadsheetApp.openById(id)
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
//   page = 1,
//   pageSize = 100,
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

const getAppUrl = () => {
  const url = ScriptApp.getService().getUrl()
  // const url = CONFIG.WebAppUrl
  return JSON.stringify(url)
}


