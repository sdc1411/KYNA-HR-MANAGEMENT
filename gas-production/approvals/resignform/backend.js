const ConfigResignForm = {
  ResponseDatabase: {
    SheetName: 'resignFormResponse',
    SheetHandover: 'resignEmployeeHandover',
    Title: 'BMHR-0202 Đơn xin thôi việc',
    DepartmentHeader: 'Phòng Ban',
    StatusHeader: '_status',
    RespondIdHeader: '_response_id',
    EmailHeader: 'Email Address',
    EmployeeHeader: 'Họ và tên nhân viên nghỉ việc',
  },
  ApprovalFlowForms: {
    ApprovalEmailForm: 'approvals/resignform/approval_email.html',
    NotificationEmailForm: 'approvals/resignform/notification_email.html',
    AprrovalIndexForm: 'approvals/resignform/index.html',
    ApprovalProgressForm: 'approvals/resignform/approval_progress.html',
  },

}


// Lấy flow duyệt từ database
function getApprovalFlowsBMHR0202() {
  const sheet = SpreadsheetApp.openById(CONFIG.DATABASE.RESIGNS);
  const data = sheet.getSheetByName(CONFIG.SHEET_NAME.RESIGN_APPROVAL_FLOW).getDataRange().getValues();

  const flows = {};

  // Skip the header row (assumed the first row contains headers)
  for (let i = 1; i < data.length; i++) {
    const [department, email, name, title] = data[i];

    if (!flows[department]) {
      flows[department] = [];
    }

    flows[department].push({ email, name, title });
  }

  return flows;
}


// lấy dữ liệu user submit form nghỉ việc từ client
const createResignForms = (params) => JSON.stringify(submittingResignForm(JSON.parse(params)))

function submittingResignForm(data) {
  
  const item = data.item
  const lastWorkingDayFormated = convertDate(item.lastWorkingDay)


  const ws = SpreadsheetApp.openById(CONFIG.DATABASE.RESIGNS)
  const sheet = ws.getSheetByName(ConfigResignForm.ResponseDatabase.SheetName)

  const responseId = Utilities.base64EncodeWebSafe(Utilities.getUuid())
  const newOffDate = ""

  if (item.resignFormType === 'other' && item.requesterEmail !== item.userEmail && (item.requesterLevel === 'Manager' || item.requesterLevel === 'Director' || item.requesterLevel === 'C-level')) {
    const dataApprovals = {"email":`${item.requesterEmail}`,"status":"Pending","comments":"","timestamp":""}
    const stringDataApprovals = JSON.stringify(dataApprovals)
    sheet.appendRow([new Date(), item.userEmail, item.employeeCode, item.fullName, item.position, item.department, item.offReason, lastWorkingDayFormated,newOffDate,responseId,'Approved',stringDataApprovals]);
    const app = new SubmitResignForm();
    app.createResignHandover();  
  } else if (item.employeeLevel === 'Manager' || item.employeeLevel === 'Director' || item.employeeLevel === 'C-level') {
    sheet.appendRow([new Date(), item.userEmail, item.employeeCode, item.fullName, item.position, item.department, item.offReason, lastWorkingDayFormated,newOffDate,responseId]);
    const app = new SubmitResignForm();
    app.onFormSubmitKeyPerson();
  } else {
    sheet.appendRow([new Date(), item.userEmail, item.employeeCode, item.fullName, item.position, item.department, item.offReason, lastWorkingDayFormated,newOffDate,responseId]);
    const app = new SubmitResignForm();
    app.onFormSubmit();
  }

  return {
    success: true,
    message: `Đơn xin thôi việc của bạn đã được gửi thành công!`,
    data: sheet,
  }

}

// flow phê duyệt đơn thôi việc
function SubmitResignForm() {

  const responseWS = SpreadsheetApp.openById(CONFIG.DATABASE.RESIGNS);
  const url = CONFIG.WebAppUrl
  const title = ConfigResignForm.ResponseDatabase.Title
  const sheetname = ConfigResignForm.ResponseDatabase.SheetName 
  const FLOWS = getApprovalFlowsBMHR0202();
  const flowHeader = ConfigResignForm.ResponseDatabase.DepartmentHeader 
  const statusHeader = ConfigResignForm.ResponseDatabase.StatusHeader
  const responseIdHeader = ConfigResignForm.ResponseDatabase.RespondIdHeader
  const emailHeader = ConfigResignForm.ResponseDatabase.EmailHeader
  const employeeHeader = ConfigResignForm.ResponseDatabase.EmployeeHeader

  const pending = "Pending"
  const approved = "Approved"
  const rejected = "Rejected"
  const waiting = "Waiting"

  const sheet = responseWS.getSheetByName(sheetname);

  this.createResignHandover = () => {
    const inputArray = sheet.getDataRange().getValues();
    const outputArray = [];

    for (let i = 1; i < inputArray.length; i++) {
      const row = inputArray[i];
      const id = row[9];
      const employeeCode = row[2];
      const employeeName = row[3];
      const position = row[4];
      const department = row[5];
      const resignReason = row[6];
      const requestResignDate = row[7];
      const resignFinalDate = row[8];
      const dataApprovals = (row[11])? JSON.parse(row[11]) : '';
      const idStatus = row[10];
      
      // console.log(dataApprovals.email)

      if (idStatus === 'Approved') {
          const resignDate = (resignFinalDate == '') ? requestResignDate : resignFinalDate
          const status = 'Pending'
          const dataHandover = (dataApprovals === '') ? '' : [{"email":`${dataApprovals.email}`,"status":"Pending","comments":"","timestamp":""},{"email":"long.nguyen@kynaforkids.vn","status":"Pending","comments":"","timestamp":""}]
          // console.log(dataHandover)
          const newRow = [
            id,
            employeeCode,
            employeeName,
            position,
            department,
            resignReason,
            resignDate,
            dataHandover,
            status
          ];
          outputArray.push(newRow);
      }
    }
    
    // console.log(outputArray)

    const destinationSS = responseWS.getSheetByName(ConfigResignForm.ResponseDatabase.SheetHandover);

    if (outputArray.length > 0) {
        for (let i = 0; i < outputArray.length; i++) {
          const entry = outputArray[i];
          
          const duplicateRow = destinationSS
            .createTextFinder(
              `${entry[0]}`
            )
            .findNext();
          
          if (!duplicateRow) {
            // Map each property to the corresponding column index in the sheet
            const row = [
              entry[0],
              entry[1],
              entry[2],
              entry[3],
              entry[4],
              entry[5],
              entry[6],
              JSON.stringify(entry[7]),
              entry[8],
            ];

            destinationSS.appendRow(row);
          }
        }
    }
  }

  function parsedValues() {
    const values = sheet.getDataRange().getDisplayValues()
    const parsedValues = values.map(value => {
      return value.map(cell => {
        try {
          return JSON.parse(cell)
        } catch (e) {
          return cell
        }
      })
    })
    return parsedValues
  }

  this.getTaskById = (id) => {
    const values = parsedValues()
    const record = values.find(value => value.some(cell => cell.taskId === id))
    const row = values.findIndex(value => value.some(cell => cell.taskId === id)) + 1

    const headers = values[0]
    const statusColumn = headers.indexOf(statusHeader) + 1
    let task
    let approver
    let nextApprover
    let column
    let approvers
    let email
    let status
    let responseId
    if (record) {
      task = record.slice(0, headers.indexOf(statusHeader) + 1).map((item, i) => {
        return {
          label: headers[i],
          value: item
        }
      })
      email = record[headers.indexOf(emailHeader)]
      status = record[headers.indexOf(statusHeader)]
      responseId = record[headers.indexOf(responseIdHeader)]
      approver = record.find(item => item.taskId === id)
      column = record.findIndex(item => item.taskId === id) + 1
      nextApprover = record[record.findIndex(item => item.taskId === id) + 1]
      approvers = record.filter(item => item.taskId)
    }
    return { email, status, responseId, task, approver, nextApprover, approvers, row, column, statusColumn }
  }

  this.getResponseById = (id) => {
    const values = parsedValues()
    const record = values.find(value => value.some(cell => cell === id))
    const headers = values[0]
    let task
    let approvers
    let status
    if (record) {
      task = record.slice(0, headers.indexOf(statusHeader) + 1).map((item, i) => {
        return {
          label: headers[i],
          value: item
        }
      })
      status = record[headers.indexOf(statusHeader)]
      approvers = record.filter(item => item.taskId)
    }
    return { task, approvers, status }
  }

  this.sendApproval = ({ task, approver, approvers }) => {
    const template = HtmlService.createTemplateFromFile(ConfigResignForm.ApprovalFlowForms.ApprovalEmailForm)

    const { responseId } = this.getTaskById(approver.taskId); // Get task data using the approver's taskId
    const values = parsedValues()
    const headers = values[0]
    const employee = values.find(value => value[headers.indexOf(responseIdHeader)] === responseId)[headers.indexOf(employeeHeader)]; // Retrieve the employee from the responses using responseId
    const department = values.find(value => value[headers.indexOf(responseIdHeader)] === responseId)[headers.indexOf(flowHeader)]; // Retrieve the department from the responses using responseId

    template.title = title
    template.task = task
    template.approver = approver
    template.approvers = approvers
    template.actionUrl = `${url}?appType=resignform&taskId=${approver.taskId}`

    template.approved = approved
    template.rejected = rejected
    template.pending = pending
    template.waiting = waiting

    const subject = "Approval Required - " + title + " - " + employee + " - " + department

    const options = {
      htmlBody: template.evaluate().getContent()
    }
    GmailApp.sendEmail(approver.email, subject, "", options)
  }


  this.sendNotification = (taskId) => {
    const { email, responseId, status, task, approvers } = this.getTaskById(taskId)
    
    const template = HtmlService.createTemplateFromFile(ConfigResignForm.ApprovalFlowForms.NotificationEmailForm)

    const values = parsedValues();
    const headers = values[0]
    const employee = values.find(value => value[headers.indexOf(responseIdHeader)] === responseId)[headers.indexOf(employeeHeader)]; // Retrieve the employee from the responses using responseId
    const department = values.find(value => value[headers.indexOf(responseIdHeader)] === responseId)[headers.indexOf(flowHeader)]; // Retrieve the department from the responses using responseId

    template.title = title
    template.task = task
    template.status = status
    template.approvers = approvers
    template.approvalProgressUrl = `${url}?appType=resignform&responseId=${responseId}`

    template.approved = approved
    template.rejected = rejected
    template.pending = pending
    template.waiting = waiting

    const subject = `Approval ${status} - ${title} - ${employee} - ${department}`

    const options = {
      htmlBody: template.evaluate().getContent()
    }
    GmailApp.sendEmail(email, subject, "", options);
  }


  this.sendCompletedNotification = (taskId) => {
    const { responseId, status, task, approvers } = this.getTaskById(taskId)
    const template = HtmlService.createTemplateFromFile(ConfigResignForm.ApprovalFlowForms.NotificationEmailForm)
    const values = parsedValues()
    const headers = values[0]
    const employee = values.find(value => value[headers.indexOf(responseIdHeader)] === responseId)[headers.indexOf(employeeHeader)]; 
    const department = values.find(value => value[headers.indexOf(responseIdHeader)] === responseId)[headers.indexOf(flowHeader)]; // Retrieve the department from the responses using responseId

    template.title = title
    template.task = task
    template.status = status
    template.approvers = approvers
    template.approvalProgressUrl = `${url}?appType=resignform&responseId=${responseId}`

    template.approved = approved
    template.rejected = rejected
    template.pending = pending
    template.waiting = waiting

    const subject = `Approval ${status} - ${title} - ${employee} - ${department}`

    const email = CONFIG.EMAIL.NotificationEmail
    const options = {
      htmlBody: template.evaluate().getContent()
    }
    GmailApp.sendEmail(email, subject, "", options)
  }

  // add addtional data to form response when update
  this.onFormSubmit = () => {
    const values = parsedValues()
    const headers = values[0]
    let lastRow = values.length
    let startColumn = headers.indexOf(statusHeader) + 1
    if (startColumn === 0) startColumn = headers.length + 1

    // const newHeaders = [statusHeader]
    const newValues = [pending]

    const flowKey = values[lastRow - 1][headers.indexOf(flowHeader)]
    const flow = FLOWS[flowKey] || FLOWS.defaultFlow
    let taskId
    flow.forEach((item, i) => {
      item.comments = null
      item.newoffdate = null
      item.taskId = Utilities.base64EncodeWebSafe(Utilities.getUuid())
      item.timestamp = new Date()
      if (i === 0) {
        item.status = pending
        taskId = item.taskId
      } else {
        item.status = waiting
      }
      if (i !== flow.length - 1) {
        item.hasNext = true
      } else {
        item.hasNext = false
      }
      newValues.push(JSON.stringify(item))
    })

    sheet.getRange(lastRow, startColumn, 1, newValues.length).setValues([newValues]);

    const { task, email, approver, approvers, nextApprover, row, column, statusColumn } = this.getTaskById(taskId)
    if (email === CONFIG.EMAIL.AutoApproveEmail) {
      approver.status = approved
      approver.timestamp = new Date()
      sheet.getRange(row, column).setValue(JSON.stringify(approver))
      sheet.getRange(row, statusColumn).setValue(approved)
      this.sendNotification(taskId)
      this.createResignHandover()
      this.sendCompletedNotification(taskId)
    } else {
      this.sendNotification(taskId)
      this.sendApproval({ task, approver, approvers })
    }

  }

  this.onFormSubmitKeyPerson = () => {
    const values = parsedValues()
    const headers = values[0]
    let lastRow = values.length
    let startColumn = headers.indexOf(statusHeader) + 1
    if (startColumn === 0) startColumn = headers.length + 1

    // const newHeaders = [statusHeader]
    const newValues = [pending]

    const flowKey = 'GroupKeyPerson'
    const flow = FLOWS[flowKey]
    let taskId
    flow.forEach((item, i) => {
      item.comments = null
      item.newoffdate = null
      item.taskId = Utilities.base64EncodeWebSafe(Utilities.getUuid())
      item.timestamp = new Date()
      if (i === 0) {
        item.status = pending
        taskId = item.taskId
      } else {
        item.status = waiting
      }
      if (i !== flow.length - 1) {
        item.hasNext = true
      } else {
        item.hasNext = false
      }
      newValues.push(JSON.stringify(item))
    })

    sheet.getRange(lastRow, startColumn, 1, newValues.length).setValues([newValues]);

    const { task, email, approver, approvers, nextApprover, row, column, statusColumn } = this.getTaskById(taskId)
    if (email === CONFIG.EMAIL.AutoApproveEmail) {
      approver.status = approved
      approver.timestamp = new Date()
      sheet.getRange(row, column).setValue(JSON.stringify(approver))
      sheet.getRange(row, statusColumn).setValue(approved)
      this.sendNotification(taskId)
      this.createResignHandover()
      this.sendCompletedNotification(taskId)
      
    } else {
      this.sendNotification(taskId)
      this.sendApproval({ task, approver, approvers })
    }

  }

  this.approve = ({ taskId, comments, newoffdate }) => {
    const { task, approver, approvers, nextApprover, row, column, statusColumn } = this.getTaskById(taskId)
    if (!approver) return
    approver.comments = comments
    approver.status = approved
    approver.timestamp = new Date()
    sheet.getRange(row, column).setValue(JSON.stringify(approver))
    if (approver.hasNext) {
      nextApprover.status = pending
      nextApprover.timestamp = new Date()
      sheet.getRange(row, column + 1).setValue(JSON.stringify(nextApprover))
      this.sendApproval({ task, approver: nextApprover, approvers })
    } else {
      sheet.getRange(row, statusColumn).setValue(approved)
      approver.newoffdate = convertDate(newoffdate)
      sheet.getRange(row, 9).setValue(approver.newoffdate)
      this.sendNotification(taskId)
      this.createResignHandover()
      this.sendCompletedNotification(taskId)
      
    }
  }

  this.reject = ({ taskId, comments, newoffdate }) => {
    const { task, approver, nextApprover, row, column, statusColumn } = this.getTaskById(taskId)
    if (!approver) return
    approver.comments = comments
    approver.newoffdate = convertDate(newoffdate)
    approver.status = rejected
    approver.timestamp = new Date()
    sheet.getRange(row, column).setValue(JSON.stringify(approver))
    sheet.getRange(row, statusColumn).setValue(rejected)
    this.sendNotification(taskId)
  }

}


function approve({ taskId, comments, newoffdate }) {
  const app = new SubmitResignForm()
  app.approve({ taskId, comments, newoffdate })
}

function reject({ taskId, comments, newoffdate }) {
  const app = new SubmitResignForm()
  app.reject({ taskId, comments, newoffdate })
}
