function getDayName(dayNumber) {
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return days[dayNumber];
}

function getWeekNumbers(dayNames) {
  const daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return dayNames.map(dayName => daysOfWeek.indexOf(dayName));
}

function getWorkingTime(startDate, endDate, startShift, endShift) {
  let startDateTime = new Date(startDate.getFullYear(), startDate.getMonth() - 1, startDate.getDate(), startShift.getHours(), startShift.getMinutes())
  let endDateTime = new Date(endDate.getFullYear(), endDate.getMonth() - 1, endDate.getDate(), endShift.getHours(), endShift.getMinutes())
  const workingMinutes = Math.abs((endDateTime.getTime() - startDateTime.getTime()) / (1000 * 60));
  const workingHours = workingMinutes / 60
  const numberWorkingDay = (workingHours <= 6 && workingHours > 2) ? 0.5 : (workingHours > 6) ? 1 : 0
  return { workingMinutes, workingHours, numberWorkingDay }
}




// Tạo lịch làm việc 
class GenerateWorkingCalendar {
  constructor() {
    this.dbEmployeeContractManagement = SpreadsheetApp.openById(CONFIG.DATABASE.EMPLOYEE_CONTRACT_MANAGEMENT)
    this.dbWorkingCalendar = SpreadsheetApp.openById(CONFIG.DATABASE.WORKING_CALENDAR)
  }

  deleteRowsBatch(database, sheetName, employeeCode, department, fromDate, toDate, batchSize) {
    const batchSizeDefault = (batchSize) ? batchSize : 100
    const startTime = new Date(fromDate)
    const endTime = new Date(toDate)
    const ss = (database === 'workingCalendarDatabase') ? this.dbWorkingCalendar : this.dbWorkingCalendar
    const sheet = ss.getSheetByName(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();
    const positionTime = headers.indexOf('startDate')
    if (employeeCode && !department) {
      const positionEmployeeCode = headers.indexOf('employeeCode')
      for (var i = data.length - 1; i >= 1; i -= batchSizeDefault) {
        var endIndex = Math.max(i - batchSizeDefault + 1, 1);

        for (var j = i; j >= endIndex; j--) {
          var rowEmployeeCode = data[j][positionEmployeeCode];
          var rowTime = new Date(data[j][positionTime]);

          if (
            rowEmployeeCode === employeeCode &&
            (rowTime.setHours(0,0,0,0) >= startTime.setHours(0,0,0,0) && rowTime.setHours(0,0,0,0) <= endTime.setHours(0,0,0,0))
          ) {
            sheet.deleteRow(j + 1);
          }
        }
      }
    } else if (department && !employeeCode) {
      const positionDepartment = headers.indexOf('department')
      for (var i = data.length - 1; i >= 1; i -= batchSizeDefault) {
        var endIndex = Math.max(i - batchSizeDefault + 1, 1);

        for (var j = i; j >= endIndex; j--) {
          var rowDepartment = data[j][positionDepartment];
          var rowTime = new Date(data[j][positionTime]);

          if (
            rowDepartment === department &&
            (rowTime.setHours(0,0,0,0) >= startTime.setHours(0,0,0,0) && rowTime.setHours(0,0,0,0) <= endTime.setHours(0,0,0,0))
          ) {
            sheet.deleteRow(j + 1);
          }
        }
      }
    } else if (department && employeeCode) {
      const positionEmployeeCode = headers.indexOf('employeeCode')
      const positionDepartment = headers.indexOf('department')
      for (var i = data.length - 1; i >= 1; i -= batchSizeDefault) {
        var endIndex = Math.max(i - batchSizeDefault + 1, 1);
        for (var j = i; j >= endIndex; j--) {
          var rowEmployeeCode = data[j][positionEmployeeCode];
          var rowDepartment = data[j][positionDepartment];
          var rowTime = new Date(data[j][positionTime]);

          if (rowEmployeeCode === employeeCode && rowDepartment === department &&
            (rowTime.setHours(0,0,0,0) >= startTime.setHours(0,0,0,0) && rowTime.setHours(0,0,0,0) <= endTime.setHours(0,0,0,0))
          ) {
            sheet.deleteRow(j + 1);
          }
        }
      }
    } else {
      for (var i = data.length - 1; i >= 1; i -= batchSizeDefault) {
        var endIndex = Math.max(i - batchSizeDefault + 1, 1);
        for (var j = i; j >= endIndex; j--) {
          var rowTime = new Date(data[j][positionTime]);
          if (
            rowTime.setHours(0,0,0,0) >= startTime.setHours(0,0,0,0) && rowTime.setHours(0,0,0,0) <= endTime.setHours(0,0,0,0)
          ) {
            sheet.deleteRow(j + 1);
          }
        }
      }
    }
  }

  generateWorkingTimeArray(fromDate, toDate, employeeCode, department) {
    const dataSheetDocTypeManagement = this.dbEmployeeContractManagement.getSheetByName(CONFIG.SHEET_NAME.DOC_TYPE_MANAGEMENT);
    const dataDocTypeManagement = dataSheetDocTypeManagement.getDataRange().getValues();
    // Create an object to store the latest contract data for each employee
    const appliedDataDocTypeManagement = [];

    for (let i = 0; i < dataDocTypeManagement.length; i++) { // Start from 1 to skip header row
      const items = {
        employeeCode: dataDocTypeManagement[i][3],
        employeeName: dataDocTypeManagement[i][4],
        department: dataDocTypeManagement[i][5],
        docType: dataDocTypeManagement[i][2],
        workingTimeType: dataDocTypeManagement[i][13],
        applyFrom: new Date(dataDocTypeManagement[i][15]),
        applyTo: new Date(dataDocTypeManagement[i][16]),
      };

      const currentDate = new Date(fromDate);
      const endDate = new Date(toDate);
      if (!items.workingTimeType || items.workingTimeType === '' || items.docType === 'OFF' || items.docType === 'QD' || items.docType === 'PLHDLD') { continue }
      if ((items.applyFrom.setHours(0,0,0,0) <= endDate.setHours(0,0,0,0)) && ((items.applyTo.setHours(0,0,0,0) >= currentDate.setHours(0,0,0,0)) || isNaN(items.applyTo))) {
        appliedDataDocTypeManagement.push(items);
      }
    }

    // console.log(appliedDataDocTypeManagement)

    const workingTimeTypesSheet = this.dbEmployeeContractManagement.getSheetByName(CONFIG.SHEET_NAME.WORKING_TIME_TYPE);
    const workingTimeTypesDataRange = workingTimeTypesSheet.getRange(1, 1, workingTimeTypesSheet.getLastRow(), workingTimeTypesSheet.getLastColumn());
    const workingTimeTypesData = workingTimeTypesDataRange.getValues();

    const data = appliedDataDocTypeManagement.map(employee => {
      const matchingWorkingTimeType = workingTimeTypesData.find(type => type[1] === employee.workingTimeType);

      if (!matchingWorkingTimeType) {
        // console.log(`No matching working time type found for employee ${employee.employeeCode}`);
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
        workingTimeTypeName: 'Lịch làm việc',
        workingTimeType: employee.workingTimeType,
        startShiftAllDay: matchingWorkingTimeType[3],
        endShiftAllDay: matchingWorkingTimeType[4],
        startShiftOptional: matchingWorkingTimeType[6],
        endShiftOptional: matchingWorkingTimeType[7],
        workingAllDay: getWeekNumbers(workingAllDay),
        workingOptionalDay: getWeekNumbers(workingOptionalDay),
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
    // console.log(employeeCode)

    const workingTimeArray = [];
    const currentDate = new Date(fromDate);
    const endDate = new Date(toDate);

    while (currentDate <= endDate) {
      const dayOfWeek = currentDate.getDay();
      const dayName = getDayName(dayOfWeek)

      for (let i = 0; i <= datafilter.length - 1; i++) {
        if (currentDate.setHours(0,0,0,0) >= datafilter[i].applyFrom.setHours(0,0,0,0) && currentDate.setHours(0,0,0,0) <= datafilter[i].applyTo.setHours(0,0,0,0) || isNaN(datafilter[i].applyTo)) {

          const workingAllDay = datafilter[i].workingAllDay || [];
          const workingOptionalDay = datafilter[i].workingOptionalDay || [];

          let startTime = "";
          let endTime = "";
          let subid = "";
          let idYear = currentDate.getFullYear()
          let idMonth = (currentDate.getMonth() + 1 < 10) ? '0' + (currentDate.getMonth() + 1) : currentDate.getMonth() + 1
          let idDate = (currentDate.getDate() < 10) ? '0' + currentDate.getDate() : currentDate.getDate()
          let idDept = datafilter[i].department.substring(0, 4) + datafilter[i].department.slice(-4)

          if (workingAllDay.includes && workingAllDay.includes(dayOfWeek)) {
            startTime = datafilter[i].startShiftAllDay;
            endTime = datafilter[i].endShiftAllDay;
            subid = "A";
            if (startTime !== "") {
              const startDate = new Date(currentDate)
              const { workingHours, numberWorkingDay } = getWorkingTime(startDate, startDate, startTime, endTime)
              const entry = {
                createdAt: new Date(),
                id: datafilter[i].employeeCode + idYear + idMonth + idDate + subid,
                idDept: idDept + idYear + idMonth + idDate,
                employeeCode: datafilter[i].employeeCode,
                employeeName: datafilter[i].employeeName,
                department: datafilter[i].department,
                docType: datafilter[i].docType,
                workingTimeType: datafilter[i].workingTimeType,
                workingTimeTypeName: datafilter[i].workingTimeTypeName,
                workingMonth: String(currentDate.getMonth() + 1),
                workingYear: String(currentDate.getFullYear()),
                workingDay: dayName,
                startShift: startTime,
                startDate: startDate,
                endShift: endTime,
                endDate: startDate,
                workingHours: workingHours,
                numberWorkingDay: numberWorkingDay,
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
              const startDate = new Date(currentDate)
              const { workingHours, numberWorkingDay } = getWorkingTime(startDate, startDate, startTime, endTime)
              const entry = {
                createdAt: new Date(),
                id: datafilter[i].employeeCode + idYear + idMonth + idDate + subid,
                idDept: idDept + idYear + idMonth + idDate,
                employeeCode: datafilter[i].employeeCode,
                employeeName: datafilter[i].employeeName,
                department: datafilter[i].department,
                docType: datafilter[i].docType,
                workingTimeType: datafilter[i].workingTimeType,
                workingTimeTypeName: datafilter[i].workingTimeTypeName,
                workingMonth: String(currentDate.getMonth() + 1),
                workingYear: String(currentDate.getFullYear()),
                workingDay: dayName,
                startShift: startTime,
                startDate: startDate,
                endShift: endTime,
                endDate: startDate,
                workingHours: workingHours,
                numberWorkingDay: numberWorkingDay,
                idStatus: 'Approved',
              };
              workingTimeArray.push(entry);
            }
          }
        }
      }
      currentDate.setDate(currentDate.getDate() + 1); // Increment currentDate by 1 day
    }
    // console.log(workingTimeArray)
    return workingTimeArray;
  }

  createWorkingCalendarReplace(fromDate, toDate, employeeCode, department, importData, importDataDetail) {
    this.deleteWorkingCalendar(fromDate, toDate, employeeCode, department, importData);
    const rows = (importData) ? importDataDetail : this.generateWorkingTimeArray(fromDate, toDate, employeeCode, department);

    const data = rows.map(entry => [
      entry.createdAt,
      entry.id,
      entry.idDept,
      entry.employeeCode,
      entry.employeeName,
      entry.department,
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
      entry.workingHours,
      entry.numberWorkingDay,
      entry.idStatus
    ]);

    let spreadsheetId = CONFIG.DATABASE.WORKING_CALENDAR
    let range = CONFIG.SHEET_NAME.WORKING_CALENDAR
    let valueInputOption = 'USER_ENTERED'

    Snippets.prototype.appendValues(spreadsheetId,range,valueInputOption,data)
  }


  createWorkingCalendarNew(fromDate, toDate, employeeCode, department, importData, importDataDetail) {
    const rows = (importData) ? importDataDetail : this.generateWorkingTimeArray(fromDate, toDate, employeeCode, department);
    const data = rows.map(entry => [
      entry.createdAt,
      entry.id,
      entry.idDept,
      entry.employeeCode,
      entry.employeeName,
      entry.department,
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
      entry.workingHours,
      entry.numberWorkingDay,
      entry.idStatus
    ]);

    let spreadsheetId = CONFIG.DATABASE.WORKING_CALENDAR
    let range = CONFIG.SHEET_NAME.WORKING_CALENDAR
    let valueInputOption = 'USER_ENTERED'

    Snippets.prototype.appendValues(spreadsheetId,range,valueInputOption,data)
  }

  deleteWorkingCalendar(fromDate, toDate, employeeCode, department, importData) {
    if (!importData) {
      this.deleteRowsBatch('workingCalendarDatabase', 'workingCalendar', employeeCode, department, fromDate, toDate)
    } else {
      importData.forEach((row) => {
        this.deleteRowsBatch('WorkingCalendarDatabase', 'workingCalendar', row.employeeCode, department, fromDate, toDate)
      })
    }
  }

}

// function testWorkingCalendar() {
//   const fromDate = '2023-10-01'
//   const toDate = '2023-10-30'
//   const employeeCode = 'K-0290'
//   const appCalendar = new GenerateWorkingCalendar()
//   appCalendar.createWorkingCalendarReplace(fromDate, toDate, employeeCode)
// }

function createWorkingCalendar(params) {
  const item = JSON.parse(params)
  const fromDate = new Date(item.fromDate)
  const toDate = new Date(item.toDate)
  const employeeCode = item.employeeCode
  const department = item.department
  const importData = (item.importData) ? item.importData : null
  const importDataDetail = (item.importDataDetail) ? item.importDataDetail : null
  const appCalendar = new GenerateWorkingCalendar()
  appCalendar.createWorkingCalendarNew(fromDate, toDate, employeeCode, department, importData, importDataDetail)
  const app = new App()
  const data = {
    database: "workingCalendarDatabase",
    sheetName: "workingCalendarRequests",
    item: {
      fromDate: fromDate,
      toDate: toDate,
      type: item.type,
      object: item.object,
      employeeCode: item.employeeCode,
      department: item.department,
      view: (!importData) ? '' : JSON.stringify(item.importData)
    }
  }
  app.createItem(data)
  
}

function updateWorkingCalendar(params) {
  const item = JSON.parse(params)
  const fromDate = new Date(item.fromDate)
  const toDate = new Date(item.toDate)
  const employeeCode = item.employeeCode
  const department = item.department
  const importData = (item.importData) ? item.importData : null
  const importDataDetail = (item.importDataDetail) ? item.importDataDetail : null
  const appCalendar = new GenerateWorkingCalendar()
  appCalendar.createWorkingCalendarReplace(fromDate, toDate, employeeCode, department, importData, importDataDetail)
  const app = new App()
  const data = {
    database: "workingCalendarDatabase",
    sheetName: "workingCalendarRequests",
    item: {
      fromDate: fromDate,
      toDate: toDate,
      type: item.type,
      object: item.object,
      employeeCode: item.employeeCode,
      department: item.department,
      view: (!importData) ? '' : JSON.stringify(item.importData)
    }
  }
  app.createItem(data)
  // const response = {
  //   success: true,
  //   message: `Lịch làm việc đang được cập nhật!`,
  // }
  // return JSON.stringify(response);
}

function deleteWorkingCalendar(params) {
  const item = JSON.parse(params)
  const fromDate = new Date(item.fromDate)
  const toDate = new Date(item.toDate)
  const employeeCode = item.employeeCode
  const department = item.department
  const importData = (item.importData) ? item.importData : null
  const appCalendar = new GenerateWorkingCalendar()
  appCalendar.deleteWorkingCalendar(fromDate, toDate, employeeCode, department, importData)
  const app = new App()
  const data = {
    database: "workingCalendarDatabase",
    sheetName: "workingCalendarRequests",
    item: {
      fromDate: fromDate,
      toDate: toDate,
      type: item.type,
      object: item.object,
      employeeCode: item.employeeCode,
      department: item.department,
      view: (!importData) ? '' : JSON.stringify(item.importData)
    }
  }
  app.createItem(data)
  // const response = {
  //   success: true,
  //   message: `Lịch làm việc đang được xóa!`,
  // }
  // return JSON.stringify(response);
}

function getWorkingCalendarData(params) {
 const inputItem = JSON.parse(params);
  const items = [];
  const databases = [
    { database: 'workingCalendarDatabase', sheetName: 'workingCalendar' },
    { database: 'timeSheetDatabase', sheetName: 'timeSheet' },
    { database: 'leaveDaysDatabase', sheetName: 'leaveTracking' },
  ];
  for (const { database, sheetName } of databases) {
    const data = app.getItems({ page: inputItem.page, pageSize: inputItem.pageSize, database: database, sheetName: sheetName, filters: inputItem.filters });
    items.push(...data.items);
  }
  const response = {
    items,
    page: inputItem.page,
  };
  return JSON.stringify(response);
}


// cập nhật lịch nghỉ phép, WFH
class PushLeaveTracking {
  processDataLeaveDay(dataLeaveDay) {
    const startShiftFormated = dataLeaveDay.startShift;
    const startShift = this.getStartShift(startShiftFormated);
    const endShiftAllDay = this.getEndShiftAllDay(startShiftFormated);
    const startDateFormated = new Date(dataLeaveDay.startDate);
    const startDate = new Date(startDateFormated.getFullYear(), startDateFormated.getMonth(), startDateFormated.getDate())
    const endShiftFormated = dataLeaveDay.endShift;
    const endShift = this.getEndShift(endShiftFormated);
    const startShiftAllDay = this.getStartShiftAllDay(endShiftFormated);
    const endDateFormated = new Date(dataLeaveDay.endDate);
    const endDate = new Date(endDateFormated.getFullYear(), endDateFormated.getMonth(), endDateFormated.getDate())
    return { startShift, endShiftAllDay, startDate, endShift, startShiftAllDay, endDate };
  }

  getStartShift(param) {
    const shiftMapping = {
      'Sáng': '08:30:00',
      'Ca A(Đầu)': '08:30:00',
      'Ca C(Đầu)': '08:30:00',
      'Ca D(Đầu)': '08:30:00',
      'Chiều': '13:00:00',
      'Ca A(Cuối)': '13:00:00',
      'Ca B(Đầu)': '14:00:00',
      'Ca B(Cuối)': '18:00:00',
      'Ca D(Cuối)': '18:00:00',
      'Ca C(Cuối)': '16:00:00'
    };
    return shiftMapping[param] || param;
  }

  getEndShiftAllDay(param) {
    const shiftMapping = {
      'Sáng': '17:30:00',
      'Chiều': '17:30:00',
      'Ca A(Đầu)': '17:30:00',
      'Ca A(Cuối)': '17:30:00',
      'Ca B(Đầu)': '22:00:00',
      'Ca B(Cuối)': '22:00:00',
      'Ca C(Đầu)': '22:00:00',
      'Ca C(Cuối)': '22:00:00',
      'Ca D(Đầu)': '22:00:00',
      'Ca D(Cuối)': '22:00:00'
    };
    return shiftMapping[param] || param;
  }

  getEndShift(param) {
    const shiftMapping = {
      'Sáng': '12:00:00',
      'Ca A(Đầu)': '12:00:00',
      'Ca D(Đầu)': '12:00:00',
      'Chiều': '17:30:00',
      'Ca A(Cuối)': '17:30:00',
      'Ca B(Đầu)': '18:00:00',
      'Ca B(Cuối)': '22:00:00',
      'Ca C(Cuối)': '22:00:00',
      'Ca D(Cuối)': '22:00:00',
      'Ca C(Đầu)': '10:30:00',
    };
    return shiftMapping[param] || param;
  }

  getStartShiftAllDay(param) {
    const shiftMapping = {
      'Sáng': '08:30:00',
      'Chiều': '08:30:00',
      'Ca A(Đầu)': '08:30:00',
      'Ca A(Cuối)': '08:30:00',
      'Ca B(Đầu)': '14:00:00',
      'Ca B(Cuối)': '14:00:00',
      'Ca C(Đầu)': '08:30:00',
      'Ca C(Cuối)': '08:30:00',
      'Ca D(Đầu)': '08:30:00',
      'Ca D(Cuối)': '08:30:00'
    };
    return shiftMapping[param] || param;
  }

  pushLeaveTracking(id) {
    const ss = SpreadsheetApp.openById(CONFIG.DATABASE.LEAVES)
    const sheet = ss.getSheetByName('leaveRequests');
    const array = sheet.getDataRange().getValues();
    const items = array.filter(item => item[1] === id && item[9] === 'Approved')[0]

    if (items.length === 0) return

    const createdAt = items[0]
    const leaveTypes = items[2]
    const employeeCode = items[3];
    const employeeName = items[4];
    const department = items[5];
    const dataLeaveDays = JSON.parse(items[7]);
    const idStatus = items[9];

    const outputArray = []

    for (const dataLeaveDay of dataLeaveDays) {
      const { startShift, endShiftAllDay, startDate, endShift, startShiftAllDay, endDate } = this.processDataLeaveDay(dataLeaveDay);
      const newRow = {
        createdAt: createdAt,
        id: id,
        employeeCode: employeeCode,
        employeeName: employeeName,
        department: department,
        leaveTypes: leaveTypes,
        startShift: startShift,
        startDate: startDate,
        endShift: endShift,
        endDate: endDate,
        startShiftAllDay: startShiftAllDay,
        endShiftAllDay: endShiftAllDay,
        idStatus: idStatus
      };
      outputArray.push(newRow);

    }

    const leaveDayArray = [];
    for (let i = 0; i < outputArray.length; i++) {
      const currentDate = new Date(outputArray[i].startDate)
      const startDate = new Date(outputArray[i].startDate);
      const endDate = new Date(outputArray[i].endDate);
      const startShift = new Date(`2023-01-01 ${outputArray[i].startShift}`)
      const startShiftAllDay = new Date(`2023-01-01 ${outputArray[i].startShiftAllDay}`)
      const endShift = new Date(`2023-01-01 ${outputArray[i].endShift}`)
      const endShiftAllDay = new Date(`2023-01-01 ${outputArray[i].endShiftAllDay}`)
      if (currentDate.setHours(0,0,0,0) === endDate.setHours(0,0,0,0)) {
        let newCurrentDate = new Date(currentDate)
        let dayOfWeek = newCurrentDate.getDay();
        let dayName = getDayName(dayOfWeek)
        let idYear = newCurrentDate.getFullYear()
        let idMonth = (newCurrentDate.getMonth() + 1 < 10) ? '0' + (newCurrentDate.getMonth() + 1) : newCurrentDate.getMonth() + 1
        let idDate = (newCurrentDate.getDate() < 10) ? '0' + newCurrentDate.getDate() : newCurrentDate.getDate()
        const idFormated = outputArray[i].employeeCode + idYear + idMonth + idDate + outputArray[i].id
        const { workingHours, numberWorkingDay } = getWorkingTime(newCurrentDate, newCurrentDate, startShift, endShift)
        const newRow = {
          createdAt: outputArray[i].createdAt,
          id: idFormated,
          employeeCode: outputArray[i].employeeCode,
          employeeName: outputArray[i].employeeName,
          department: outputArray[i].department,
          leaveTypes: outputArray[i].leaveTypes,
          workingMonth: String(newCurrentDate.getMonth() + 1),
          workingYear: String(newCurrentDate.getFullYear()),
          workingDay: dayName,
          startShift: outputArray[i].startShift,
          startDate: newCurrentDate,
          endShift: outputArray[i].endShift,
          endDate: newCurrentDate,
          workingHours: workingHours,
          numberWorkingDay: numberWorkingDay,
          idStatus: outputArray[i].idStatus
        }
        leaveDayArray.push(newRow)
      }
      if (currentDate.setHours(0,0,0,0) < endDate.setHours(0,0,0,0)) {
        while (currentDate <= endDate) {
          let newCurrentDate = new Date(currentDate)
          let dayOfWeek = newCurrentDate.getDay();
          let dayName = getDayName(dayOfWeek)
          let idYear = newCurrentDate.getFullYear()
          let idMonth = (newCurrentDate.getMonth() + 1 < 10) ? '0' + (newCurrentDate.getMonth() + 1) : newCurrentDate.getMonth() + 1
          let idDate = (newCurrentDate.getDate() < 10) ? '0' + newCurrentDate.getDate() : newCurrentDate.getDate()
          let idFormated = outputArray[i].employeeCode + idYear + idMonth + idDate + outputArray[i].id
          if (currentDate.setHours(0,0,0,0) === startDate.setHours(0,0,0,0)) {
            const { workingHours, numberWorkingDay } = getWorkingTime(newCurrentDate, newCurrentDate, startShift, endShiftAllDay)
            const newRow = {
              createdAt: outputArray[i].createdAt,
              id: idFormated,
              employeeCode: outputArray[i].employeeCode,
              employeeName: outputArray[i].employeeName,
              department: outputArray[i].department,
              leaveTypes: outputArray[i].leaveTypes,
              workingMonth: String(newCurrentDate.getMonth() + 1),
              workingYear: String(newCurrentDate.getFullYear()),
              workingDay: dayName,
              startShift: outputArray[i].startShift,
              startDate: newCurrentDate,
              endShift: outputArray[i].endShiftAllDay,
              endDate: newCurrentDate,
              workingHours: workingHours,
              numberWorkingDay: numberWorkingDay,
              idStatus: outputArray[i].idStatus
            }
            leaveDayArray.push(newRow)
          }
          if (currentDate.setHours(0,0,0,0) > startDate.setHours(0,0,0,0) && currentDate.setHours(0,0,0,0) < endDate.setHours(0,0,0,0)) {
            const { workingHours, numberWorkingDay } = getWorkingTime(newCurrentDate, newCurrentDate, startShiftAllDay, endShiftAllDay)
            const newRow = {
              createdAt: outputArray[i].createdAt,
              id: idFormated,
              employeeCode: outputArray[i].employeeCode,
              employeeName: outputArray[i].employeeName,
              department: outputArray[i].department,
              leaveTypes: outputArray[i].leaveTypes,
              workingMonth: String(newCurrentDate.getMonth() + 1),
              workingYear: String(newCurrentDate.getFullYear()),
              workingDay: dayName,
              startShift: outputArray[i].startShiftAllDay,
              startDate: newCurrentDate,
              endShift: outputArray[i].endShiftAllDay,
              endDate: newCurrentDate,
              workingHours: workingHours,
              numberWorkingDay: numberWorkingDay,
              idStatus: outputArray[i].idStatus
            }
            leaveDayArray.push(newRow)
          }
          if (currentDate.setHours(0,0,0,0) === endDate.setHours(0,0,0,0)) {
            const { workingHours, numberWorkingDay } = getWorkingTime(newCurrentDate, newCurrentDate, startShiftAllDay, endShift)
            const newRow = {
              createdAt: outputArray[i].createdAt,
              id: idFormated,
              employeeCode: outputArray[i].employeeCode,
              employeeName: outputArray[i].employeeName,
              department: outputArray[i].department,
              leaveTypes: outputArray[i].leaveTypes,
              workingMonth: String(newCurrentDate.getMonth() + 1),
              workingYear: String(newCurrentDate.getFullYear()),
              workingDay: dayName,
              startShift: outputArray[i].startShiftAllDay,
              startDate: newCurrentDate,
              endShift: outputArray[i].endShift,
              endDate: newCurrentDate,
              workingHours: workingHours,
              numberWorkingDay: numberWorkingDay,
              idStatus: outputArray[i].idStatus
            }
            leaveDayArray.push(newRow)
          }
          currentDate.setDate(currentDate.getDate() + 1); // Increment currentDate by 1 day
        }
      }
    }

    const destinationSS = ss.getSheetByName('leaveTracking');
    if (leaveDayArray.length > 0) {
      const existingIds = new Set(destinationSS.getRange(1, 2, destinationSS.getLastRow(), 1).getValues().flat());

      const results = leaveDayArray
        .filter(entry => !existingIds.has(entry.id))
        .map(entry => [
          entry.createdAt,
          entry.id,
          entry.employeeCode,
          entry.employeeName,
          entry.department,
          entry.leaveTypes,
          entry.workingMonth,
          entry.workingYear,
          entry.workingDay,
          entry.startShift,
          entry.startDate,
          entry.endShift,
          entry.endDate,
          entry.workingHours,
          entry.numberWorkingDay,
          entry.idStatus
        ]);

      if (results.length > 0) {
        destinationSS.getRange(destinationSS.getLastRow() + 1, 1, results.length, results[0].length).setValues(results);
      }
    }
  }

}

// function testGetValueNormal () {
//   const ws = SpreadsheetApp.openById(CONFIG.DATABASE.WORKING_CALENDAR)
//   const ss = ws.getSheetByName(CONFIG.SHEET_NAME.WORKING_CALENDAR)
//   const data = ss.getDataRange().getValues()
//   console.log(data)
// }

// function testGetValueAPI () {
//   const data = Snippets.prototype.getValues(CONFIG.DATABASE.WORKING_CALENDAR,CONFIG.SHEET_NAME.WORKING_CALENDAR).values
//   console.log(data)
// }
