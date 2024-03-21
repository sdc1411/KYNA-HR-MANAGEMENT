function getTimeSheet(params) {
  const item = JSON.parse(params)
  const fromDate = item.fromDate
  const toDate = item.toDate
  const top = item.top
  const otherCondition = item.otherCondition
  const server = readSingleProperty('serverTimeSheet');
  const username = readSingleProperty('usernameTimeSheet');
  const password = readSingleProperty('passwordTimeSheet');
  const db = readSingleProperty('databaseTimeSheet')

  const queryStartDateFormated = "'" + fromDate + "'"
  const queryEndDateFormated = "'" + toDate + "'"
  const topFormated = (top) ? `top ${top} ` : ''
  const otherConditionFormated = (otherCondition) ? ` and ${otherCondition}` : ''

  const queryString = "SELECT " + topFormated + "('K-' + RIGHT('0000' + CAST(UserEnrollNumber AS VARCHAR(4)), 4)) as employeeCode, TimeDate as startDate, TimeDate as endDate, min(TimeStr) as startShift, max(TimeStr) as endShift FROM dbo.CheckInOut where TimeDate >= " + queryStartDateFormated + " and TimeDate <= " + queryEndDateFormated + otherConditionFormated + " group by UserEnrollNumber, TimeDate"
  const dbUrl = 'jdbc:sqlserver://' + server + ':1433;databaseName=' + db;
  const databaseUsers = SpreadsheetApp.openById(CONFIG.DATABASE.EMPLOYEE_INFORMATION).getSheetByName(CONFIG.SHEET_NAME.RAW_DATA_USERS)
  const listUsers = databaseUsers.getDataRange().getValues()
  const createdAt = new Date()

  const rowData = [];
  let conn = null
  try {
    conn = Jdbc.getConnection(dbUrl, username, password);
    const stmt = conn.createStatement();
    const results = stmt.executeQuery(queryString);

    while (results.next()) {
      const employeeCode = results.getString('employeeCode');
      const department = listUsers.filter(item => item[0] === employeeCode).map(item => item[63])
      const startDate = new Date(results.getString('startDate'));
      const endDate = new Date(results.getString('endDate'));
      const startShift = new Date(results.getString('startShift'));
      const endShift = new Date(results.getString('endShift'));
      const startShiftFormat = Utilities.formatDate(startShift, Session.getScriptTimeZone(), 'h:mm:ss a')
      const endShiftFormat = Utilities.formatDate(endShift, Session.getScriptTimeZone(), 'h:mm:ss a')
      const workingMonth = String(startDate.getMonth() + 1)
      const workingYear = String(startDate.getFullYear())
      const workingDay = getDayName(startDate.getDay())
      const {workingHours, numberWorkingDay} = getWorkingTime(startDate,endDate,startShift,endShift)
      const numberWorkingDayFormat = (workingHours === 0) ? 1 : numberWorkingDay
      let idYear = startDate.getFullYear()
      let idMonth = (startDate.getMonth() + 1 < 10) ? '0' + (startDate.getMonth() + 1) : startDate.getMonth() + 1
      let idDate = (startDate.getDate() < 10) ? '0' + startDate.getDate() : startDate.getDate()
      const idFormated = employeeCode + idYear + idMonth + idDate + 'VT'

      rowData.push([createdAt,idFormated, employeeCode, department[0],'Chấm công vân tay', workingMonth,workingYear,workingDay, startShiftFormat, startDate, endShiftFormat, endDate, workingHours,numberWorkingDay, 'Approved']);
    }
    results.close();
    stmt.close();
    conn.close();
  } catch (e) {
    console.error('Error: ' + e);
  } finally {
    if (conn) {
      conn.close()
    }
  }

  if (rowData.length > 0) {
    let spreadsheetId = CONFIG.DATABASE.TIME_SHEET
    let range = CONFIG.SHEET_NAME.TIME_SHEET
    let valueInputOption = 'USER_ENTERED'
    Snippets.prototype.appendValues(spreadsheetId,range,valueInputOption,rowData)
    const app = new App()
    const data = {
      database: "timeSheetDatabase",
      sheetName: "timeSheetRequests",
      item: {
        fromDate: new Date(fromDate),
        toDate: new Date(toDate),
        type: item.type,
        top: top,
        otherCondition: otherCondition
      }
    }
    app.createItem(data)
  } else {
    return console.error('Không có dữ liệu')
  }

}

// function testTimeSheet () {

//   const fromDate = '2023-11-01'
//   const toDate = '2023-11-01'
//   const top = null
//   const otherCondition = null
//   const server = readSingleProperty('serverTimeSheet');
//   const username = readSingleProperty('usernameTimeSheet');
//   const password = readSingleProperty('passwordTimeSheet');
//   const db = readSingleProperty('databaseTimeSheet')

//   const queryStartDateFormated = "'" + fromDate + "'"
//   const queryEndDateFormated = "'" + toDate + "'"
//   const topFormated = (top) ? `top ${top} ` : ''
//   const otherConditionFormated = (otherCondition) ? ` and ${otherCondition}` : ''

//   const queryString = "SELECT " + topFormated + "('K-' + RIGHT('0000' + CAST(UserEnrollNumber AS VARCHAR(4)), 4)) as employeeCode, TimeDate as startDate, TimeDate as endDate, min(TimeStr) as startShift, max(TimeStr) as endShift FROM dbo.CheckInOut where TimeDate >= " + queryStartDateFormated + " and TimeDate <= " + queryEndDateFormated + otherConditionFormated + " group by UserEnrollNumber, TimeDate"
//   const dbUrl = 'jdbc:sqlserver://' + server + ':1433;databaseName=' + db;
//   const databaseUsers = SpreadsheetApp.openById(CONFIG.DATABASE.EMPLOYEE_INFORMATION).getSheetByName(CONFIG.SHEET_NAME.RAW_DATA_USERS)
//   const listUsers = databaseUsers.getDataRange().getValues()
//   const createdAt = new Date()

//   const rowData = [];
//   let conn = null
//   try {
//     conn = Jdbc.getConnection(dbUrl, username, password);
//     const stmt = conn.createStatement();
//     const results = stmt.executeQuery(queryString);

//     while (results.next()) {
//       const employeeCode = results.getString('employeeCode');
//       const department = listUsers.filter(item => item[0] === employeeCode).map(item => item[63])
//       const startDate = new Date(results.getString('startDate'));
//       const endDate = new Date(results.getString('endDate'));
//       const startShift = new Date(results.getString('startShift'));
//       const endShift = new Date(results.getString('endShift'));
//       const startShiftFormat = Utilities.formatDate(startShift, Session.getScriptTimeZone(), 'h:mm:ss a')
//       const endShiftFormat = Utilities.formatDate(endShift, Session.getScriptTimeZone(), 'h:mm:ss a')
//       const workingMonth = String(startDate.getMonth() + 1)
//       const workingYear = String(startDate.getFullYear())
//       const workingDay = getDayName(startDate.getDay())
//       const {workingHours, numberWorkingDay} = getWorkingTime(startDate,endDate,startShift,endShift)
//       let idYear = startDate.getFullYear()
//       let idMonth = (startDate.getMonth() + 1 < 10) ? '0' + (startDate.getMonth() + 1) : startDate.getMonth() + 1
//       let idDate = (startDate.getDate() < 10) ? '0' + startDate.getDate() : startDate.getDate()
//       const idFormated = employeeCode + idYear + idMonth + idDate + 'VT'

//       rowData.push([createdAt,idFormated, employeeCode, department[0],'Chấm công vân tay', workingMonth,workingYear,workingDay, startShiftFormat, startDate, endShiftFormat, endDate, workingHours,numberWorkingDay, 'Approved']);
//     }
//     results.close();
//     stmt.close();
//     conn.close();
//   } catch (e) {
//     console.error('Error: ' + e);
//   } finally {
//     if (conn) {
//       conn.close()
//     }
//   }
//   console.log(rowData)

// }
