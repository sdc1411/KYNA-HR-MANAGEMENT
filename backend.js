const CONFIG = {
  WebAppUrl: "https://script.google.com/macros/s/AKfycbw90tw50FCBUTqENy3cW86u_dJgxuIz_rufRVz72r32UwDI6vnCZuzrqGemkIL5gIKVVg/exec",
  INDEX: "index.html",
  NAME: "KYNA HR DASHBOARD",
  TITLE: "Welcome to HR Dashboard",
  PAGE_SIZE: 15,
  REVERSE: true,
  DB_USER: '1U1US3obVZMpPquEEalr8lpHwlFm4u9KGWvhs45kURSI',
  SHEET_NAME_USER: 'users',
  DB_LEAVES: '1_5H7EbRGjpOW5NkB6UfNSg_r8QnHOJiXsXAFkje0ZbA',
  SHEET_NAME_LEAVES: 'leaveTypes',
  DB_WORKING_CALENDAR: '1tG6S-2wHwEBEi6IvV1a0oxyHjDKR1VKLPA6M-Y7XZG8',
  SHEET_WORKING_CALENDAR: 'workingCalendar',
  DB_EMPLOYEE_INFORMATION: '1K4TtcaK0hrRa0yGv0xSI2jP4WI8WjhGrE-pgPdcx0Mk',
  SHEET_UPDATE_PROFILE_REQUEST: 'updateProfileRequests',
  STATUS: {
    APPROVED: "Approved",
    REJECTED: "Rejected",
  }
}

// chức năng convert ngày tháng từ Vue và từ Boostrap
function convertDate(data) {
  if (data.includes("-")) {
    const dateParts1 = data.split("-")
    const dateFormated1 = new Date(+dateParts1[0], dateParts1[1] - 1, +dateParts1[2]);
    console.log(dateFormated1)
    return dateFormated1
    
  } else {
    const dateParts2 = data.split("/")
    const dateFormated2 = new Date(+dateParts2[2], dateParts2[1] - 1, +dateParts2[0]);
    console.log(dateFormated2)
    return dateFormated2
  }
}


function doGet(e) {

  const appType = e.parameters.appType;

  if (appType === 'reviewform') {
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

  } else if (appType === 'resignform') {
    
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
          console.log(template.task)
        } else if (e.parameter.responseId) {
          template = HtmlService.createTemplateFromFile(ConfigResignForm.ApprovalFlowForms.ApprovalProgressForm);
          const { task, approvers, status } = appResignForm.getResponseById(e.parameter.responseId);
          template.task = task;
          template.status = status;
          template.approvers = approvers;
          console.log(template.task)
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
}





function include_(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent()
}



class App {
  constructor() {
    this.dbUser = SpreadsheetApp.openById(CONFIG.DB_USER)
    this.dbLeaveType = SpreadsheetApp.openById(CONFIG.DB_LEAVES)
    this.reverse = CONFIG.REVERSE
    this.headerId = 'id'
  }

  getAppInfo() {
    const data = {
      name: CONFIG.NAME,
      title: CONFIG.TITLE,
    }
    return data
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
    const ws = this.dbUser.getSheetByName(CONFIG.SHEET_NAME_USER)
    if (!ws) throw new Error(`${CONFIG.SHEET_NAME_USER} was not found in the database`)
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
    const ws = this.dbUser.getSheetByName(CONFIG.SHEET_NAME_USER)
    if (!ws) throw new Error(`${CONFIG.SHEET_NAME_USER} was not found in the database`)
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


  // getListLeaveType () {
  //   const ws = this.dbLeaveType.getSheetByName(CONFIG.SHEET_NAME_LEAVES)
  //   if (!ws) throw new Error(`${CONFIG.SHEET_NAME_LEAVES} was not found in the database`)
  //   const [headers,...records] = ws.getDataRange().getValues()
    
  //   const keys = this.createKeys(headers)

  //   const leaveTypeArray = records.map(values => this.createItemObject(keys, values))

  //   return leaveTypeArray
  // }

  getItems({ page, pageSize, getItemDB ,sheetName, filters }) {
    const workSheetID = (getItemDB === 'userDatabase') ? CONFIG.DB_USER : (getItemDB === 'leaveDaysDatabase') ? CONFIG.DB_LEAVES : (getItemDB === 'workingCalendarDatabase') ? CONFIG.DB_WORKING_CALENDAR : null
    const ss = SpreadsheetApp.openById(workSheetID)
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
    return {
      pages: Math.ceil(records.length / pageSize),
      items: records.slice((page - 1) * pageSize, (page - 1) * pageSize + pageSize).map(record => this.createItemObject(keys, record)),
    }
    
  }

  createItem({ createItemDB ,sheetName, item }) {
    const workSheetID = (createItemDB === 'leaveDaysDatabase') ? CONFIG.DB_LEAVES : (createItemDB === 'employeeInformationDatabase') ? CONFIG.DB_EMPLOYEE_INFORMATION : null
    const ss = SpreadsheetApp.openById(workSheetID)
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

  updateApproval({ updateApprovalDB ,sheetName, item }) {
    const workSheetID = (updateApprovalDB === 'leaveDaysDatabase') ? CONFIG.DB_LEAVES : null
    const ss = SpreadsheetApp.openById(workSheetID)
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
    console.log(values)
    console.log(keys)
    console.log(item)
    return {
      success: true,
      message: `Item ${item.id} has been updated successfully!`,
      data: item,
    }
  }

}

const app = new App()

const test = (params) => {
  
  
  const data = app.getListLeaveType()
  console.log(data)
  const detail = data.find(item => item['leaveTypeCode'] === 'NO')
  console.log(detail)
  const lyDoList = data.map(item => item['leaveTypeName']);
  console.log(lyDoList)
  const details2 = detail.leaveTypeName
  console.log(details2)
  return JSON.stringify(data)
  
}


const getDataLoginUser = (params) => JSON.stringify(app.getDataLoginUser(params))
const getDataLoginNonUser = (params) => JSON.stringify(app.getDataLoginNonUser(params))

const getAppInfo = () => JSON.stringify(app.getAppInfo())

const getItems = (params) => JSON.stringify(app.getItems(JSON.parse(params)))
const createItem = (params) => JSON.stringify(app.createItem(JSON.parse(params)))
const updateApproval = (params) => JSON.stringify(app.updateApproval(JSON.parse(params)))

