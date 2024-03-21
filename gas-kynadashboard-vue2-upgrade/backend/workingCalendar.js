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

  generateWorkingTimeArray(fromDate, toDate, employeeCode, department, importData) {
    const dataSheetDocTypeManagement = this.dbEmployeeContractManagement.getSheetByName(CONFIG.SHEET_NAME.DOC_TYPE_MANAGEMENT);
    const dataDocTypeManagement = dataSheetDocTypeManagement.getDataRange().getValues();
    // Create an object to store the latest contract data for each employee
    const appliedDataDocTypeManagement = [];

    if (!importData) {
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

        const currentDate = new Date(fromDate);
        const endDate = new Date(toDate);
        if (!items.workingTimeType || items.workingTimeType === '' || items.docType === 'OFF' || items.docType === 'QD' || items.docType === 'PLHDLD') { continue }
        if (items.applyFrom <= endDate && (items.applyTo >= currentDate || items.applyTo === '')) {
          appliedDataDocTypeManagement.push(items);
        }
      }
    } else {
      for (let i = 0; i < importData.length; i++) {
        const items = {
          employeeCode: importData[i][0],
          employeeName: importData[i][1],
          department: department,
          docType: "",
          workingTimeType: importData[i][2],
          applyFrom: "",
          applyTo: "",
        }
        appliedDataDocTypeManagement.push(items);
      }
    }

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
        workingAllDay: this.getWeekNumbers(workingAllDay),
        workingOptionalDay: this.getWeekNumbers(workingOptionalDay),
      };
    }).filter(entry => entry !== null);

    const datafilter = data.filter(item =>
      (department && employeeCode) ?
        item.department === department && item.employeeCode === employeeCode :
        (department && !employeeCode) ?
          item.department === department :
          (!department && employeeCode) ?
            item.employeeCode === employeeCode :
            true
    );

    const workingTimeArray = [];
    const currentDate = new Date(fromDate);
    const endDate = new Date(toDate);

    while (currentDate <= endDate) {
      const dayOfWeek = currentDate.getDay();
      const dayName = this.getDayName(dayOfWeek)

      for (let i = 0; i <= datafilter.length - 1; i++) {
        if (currentDate >= datafilter[i].applyFrom && currentDate <= datafilter[i].applyTo || datafilter[i].applyTo === '') {

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
              const entry = {
                createAt: new Date(),
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
                startDate: new Date(currentDate),
                endShift: endTime,
                endDate: new Date(currentDate),
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
                startDate: new Date(currentDate),
                endShift: endTime,
                endDate: new Date(currentDate),
                idStatus: 'Approved',
              };
              workingTimeArray.push(entry);
            }
          }
        }
      }
      currentDate.setDate(currentDate.getDate() + 1); // Increment currentDate by 1 day
    }
    return workingTimeArray;
  }

  createWorkingCalendarReplace(fromDate, toDate, employeeCode, department, importData) {
    const destinationSS = this.dbWorkingCalendar.getSheetByName('workingCalendar')
    const rows = this.generateWorkingTimeArray(fromDate, toDate, employeeCode, department, importData);
    this.deleteWorkingCalendar(fromDate, toDate, employeeCode, department, importData)

    const data = rows.map(entry => [
      entry.createAt,
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
      entry.idStatus
    ]);
    const destinationValue = destinationSS.getDataRange()
    const lastRow = destinationValue.getLastRow()
    const lastCol = destinationValue.getLastColumn()
    destinationSS.getRange(lastRow + 1, 1, rows.length, lastCol).setValues(data)

  }

  createWorkingCalendarNew(fromDate, toDate, employeeCode, department) {
    const destinationSS = this.dbWorkingCalendar.getSheetByName('workingCalendar')
    const rows = this.generateWorkingTimeArray(fromDate, toDate, employeeCode, department);

    const data = rows.map(entry => [
      entry.createAt,
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
      entry.idStatus
    ]);
    const destinationValue = destinationSS.getDataRange()
    const lastRow = destinationValue.getLastRow()
    const lastCol = destinationValue.getLastColumn()
    destinationSS.getRange(lastRow + 1, 1, rows.length, lastCol).setValues(data)
  }

  deleteWorkingCalendar(fromDate, toDate, employeeCode, department, importData) {
    const destinationSS = this.dbWorkingCalendar.getSheetByName('workingCalendar')
    const currentDate = new Date(fromDate);
    const endDate = new Date(toDate);
    if (!importData) {
      const deletedWorkingTimeArray = [];
      while (currentDate <= endDate) {
        let idYear = currentDate.getFullYear()
        let idMonth = (currentDate.getMonth() + 1 < 10) ? '0' + (currentDate.getMonth() + 1) : currentDate.getMonth() + 1
        let idDate = (currentDate.getDate() < 10) ? '0' + currentDate.getDate() : currentDate.getDate()
        let idEmp = (employeeCode) ? employeeCode : ''
        let idDept = (department) ? department.substring(0, 4) + department.slice(-4) : ''

        const entryA = {
          id: idEmp + idYear + idMonth + idDate + "A",
          idDept: idDept + idYear + idMonth + idDate,
        };
        deletedWorkingTimeArray.push(entryA);

        const entryB = {
          id: idEmp + idYear + idMonth + idDate + "B",
          idDept: idDept + idYear + idMonth + idDate,
        };
        deletedWorkingTimeArray.push(entryB);
        currentDate.setDate(currentDate.getDate() + 1); // Increment currentDate by 1 day
      }
      if (employeeCode) {
        for (const entry of deletedWorkingTimeArray) {
          let matchingRow;
          while ((matchingRow = destinationSS.createTextFinder(entry.id).findNext())) {
            // Get the row number of the matching entry
            const rowIndex = matchingRow.getRow();
            // Delete the entire row
            destinationSS.deleteRow(rowIndex);
          }
        }
      } else if (department) {
        for (const entry of deletedWorkingTimeArray) {
          let matchingRow;
          while ((matchingRow = destinationSS.createTextFinder(entry.idDept).findNext())) {
            // Get the row number of the matching entry
            const rowIndex = matchingRow.getRow();
            // Delete the entire row
            destinationSS.deleteRow(rowIndex);
          }
        }
      } else return

    } else {
      const deletedWorkingTimeArray = [];
      while (currentDate <= endDate) {
        for (let i = 0; i < importData.length; i++) {
          let idYear = currentDate.getFullYear()
          let idMonth = (currentDate.getMonth() + 1 < 10) ? '0' + (currentDate.getMonth() + 1) : currentDate.getMonth() + 1
          let idDate = (currentDate.getDate() < 10) ? '0' + currentDate.getDate() : currentDate.getDate()
          let idEmp = (importData[i][0]) ? importData[i][0] : ''
          let idDept = (department) ? department.substring(0, 4) + department.slice(-4) : ''
          const entryA = {
            id: idEmp + idYear + idMonth + idDate + "A",
            idDept: idDept + idYear + idMonth + idDate,
          };
          deletedWorkingTimeArray.push(entryA);

          const entryB = {
            id: idEmp + idYear + idMonth + idDate + "B",
            idDept: idDept + idYear + idMonth + idDate,
          };
          deletedWorkingTimeArray.push(entryB);
        }
        currentDate.setDate(currentDate.getDate() + 1); // Increment currentDate by 1 day
      }
      for (const entry of deletedWorkingTimeArray) {
        let matchingRow;
        while ((matchingRow = destinationSS.createTextFinder(entry.id).findNext())) {
          // Get the row number of the matching entry
          const rowIndex = matchingRow.getRow();
          // Delete the entire row
          destinationSS.deleteRow(rowIndex);
        }
      }

    }
  }
}


function createWorkingCalendar(params) {
  const item = JSON.parse(params)
  const object = item.object
  const fromDate = new Date(item.fromDate)
  const toDate = new Date(item.toDate)
  const employeeCode = item.employeeCode
  const department = item.department
  if (object === 'Tất cả phòng ban') {
    const appCalendar = new GenerateWorkingCalendar()
    appCalendar.createWorkingCalendarNew(fromDate, toDate, employeeCode, department)
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
        department: item.department
      }
    }
    app.createItem(data)

    // const response = {
    //   success: true,
    //   message: `Lịch làm việc đang được tạo!`,
    // }
    // return JSON.stringify(response);
  } else {
    const appCalendar = new GenerateWorkingCalendar()
    appCalendar.createWorkingCalendarReplace(fromDate, toDate, employeeCode, department)
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
        department: item.department
      }
    }
    app.createItem(data)

    // const response = {
    //   success: true,
    //   message: `Lịch làm việc đang được tạo!`,
    // }
    // return JSON.stringify(response);
  }
}

function updateWorkingCalendar(params) {
  const item = JSON.parse(params)
  const fromDate = new Date(item.fromDate)
  const toDate = new Date(item.toDate)
  const employeeCode = item.employeeCode
  const department = item.department
  const importData = (item.importData) ? item.importData : null
  const appCalendar = new GenerateWorkingCalendar()
  appCalendar.createWorkingCalendarReplace(fromDate, toDate, employeeCode, department, importData)
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
      department: item.department
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
      department: item.department
    }
  }
  app.createItem(data)
  // const response = {
  //   success: true,
  //   message: `Lịch làm việc đang được xóa!`,
  // }
  // return JSON.stringify(response);
}


// cập nhật lịch nghỉ phép, WFH
function pushLeaveTracking() {
  const ss = SpreadsheetApp.openById(CONFIG.DATABASE.LEAVES)
  const sheet = ss.getSheetByName('leaveRequests');
  const inputArray = sheet.getDataRange().getValues();
  const outputArray = [];

  for (let i = 1; i < inputArray.length; i++) {
    const row = inputArray[i];
    const createdAt = row[0];
    const id = row[1];
    const leaveTypes = row[2];
    const employeeCode = row[3];
    const employeeName = row[4];
    const department = row[5];
    const dataLeaveDays = JSON.parse(row[7]);
    const idStatus = row[9];

    if (idStatus === 'Approved') {
      for (const dataLeaveDay of dataLeaveDays) {
        const startShiftFormated = dataLeaveDay.startShift;
        const startShift = (startShiftFormated === 'Sáng' || startShiftFormated === 'Ca A(Đầu)' || startShiftFormated === 'Ca C(Đầu)' || startShiftFormated === 'Ca D(Đầu)') ? '08:30:00' : (startShiftFormated === 'Chiều' || startShiftFormated === 'Ca A(Cuối)') ? '13:00:00' : (startShiftFormated === 'Ca B(Đầu)') ? '14:00:00' : (startShiftFormated === 'Ca B(Cuối)') ? '18:00:00' : (startShiftFormated === 'Ca C(Cuối)') ? '18:00:00' : (startShiftFormated === 'Ca D(Cuối)') ? '16:00:00' : startShiftFormated
        const startDateFormated = new Date(dataLeaveDay.startDate);
        const startDate = new Date(startDateFormated.getFullYear(), startDateFormated.getMonth(), startDateFormated.getDate())
        const endShiftFormated = dataLeaveDay.endShift;
        const endShift = (endShiftFormated === 'Sáng' || endShiftFormated === 'Ca A(Đầu)' || endShiftFormated === 'Ca C(Đầu)') ? '12:00:00' : (endShiftFormated === 'Chiều' || endShiftFormated === 'Ca A(Cuối)') ? '17:30:00' : (endShiftFormated === 'Ca B(Đầu)') ? '18:00:00' : (endShiftFormated === 'Ca B(Cuối)') ? '22:00:00' : (endShiftFormated === 'Ca C(Cuối)' || endShiftFormated === 'Ca D(Cuối)') ? '22:00:00' : (endShiftFormated === 'Ca D(Đầu)') ? '10:30:00' : endShiftFormated
        const endDateFormated = new Date(dataLeaveDay.endDate);
        const endDate = new Date(endDateFormated.getFullYear(), endDateFormated.getMonth(), endDateFormated.getDate())
        let idYear = startDateFormated.getFullYear()
        let idMonth = (startDateFormated.getMonth() + 1 < 10) ? '0' + (startDateFormated.getMonth() + 1) : startDateFormated.getMonth() + 1
        let idDate = (startDateFormated.getDate() < 10) ? '0' + startDateFormated.getDate() : startDateFormated.getDate()
        const idFormated = employeeCode + idYear + idMonth + idDate + id

        const newRow = [
          createdAt,
          idFormated,
          employeeCode,
          employeeName,
          department,
          leaveTypes,
          startShift,
          startDate,
          endShift,
          endDate,
          idStatus
        ];
        outputArray.push(newRow);

      }
    }
  }

  const destinationSS = ss.getSheetByName('leaveTracking');

  if (outputArray.length > 0) {
    for (let i = 0; i < outputArray.length; i++) {
      const entry = outputArray[i];

      const duplicateRow = destinationSS
        .createTextFinder(
          `${entry[1]}`
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
          entry[7],
          entry[8],
          entry[9],
          entry[10]
        ];

        destinationSS.appendRow(row);
      }
    }
  }
}
