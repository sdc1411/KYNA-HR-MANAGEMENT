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
    const dataLeaveDays = JSON.parse(row[7]);
    const idStatus = row[9];

    if (idStatus === 'Approved') {
      for (const dataLeaveDay of dataLeaveDays) {
        const startShiftFormated = dataLeaveDay.startShift;
        const startShift = (startShiftFormated === 'Sáng' || startShiftFormated === 'Ca A(Đầu)' || startShiftFormated === 'Ca C(Đầu)' || startShiftFormated === 'Ca D(Đầu)' ) ? '08:30:00' : (startShiftFormated === 'Chiều' || startShiftFormated === 'Ca A(Cuối)') ? '13:00:00' : (startShiftFormated === 'Ca B(Đầu)') ? '14:00:00' : (startShiftFormated === 'Ca B(Cuối)') ? '18:00:00' : (startShiftFormated === 'Ca C(Cuối)') ? '18:00:00' : (startShiftFormated === 'Ca D(Cuối)') ? '16:00:00' : startShiftFormated
        const startDateFormated = new Date(dataLeaveDay.startDate);
        const startDate = new Date(startDateFormated.getFullYear(),startDateFormated.getMonth(),startDateFormated.getDate())
        const endShiftFormated = dataLeaveDay.endShift;
        const endShift = (endShiftFormated === 'Sáng' || endShiftFormated === 'Ca A(Đầu)' || endShiftFormated === 'Ca C(Đầu)' ) ? '12:00:00' : (endShiftFormated === 'Chiều' || endShiftFormated === 'Ca A(Cuối)') ? '17:30:00' : (endShiftFormated === 'Ca B(Đầu)') ? '18:00:00' : (endShiftFormated === 'Ca B(Cuối)') ? '22:00:00' : (endShiftFormated === 'Ca C(Cuối)' || endShiftFormated === 'Ca D(Cuối)') ? '22:00:00' : (endShiftFormated === 'Ca D(Đầu)') ? '10:30:00' : endShiftFormated
        const endDateFormated = new Date(dataLeaveDay.endDate);
        const endDate = new Date(endDateFormated.getFullYear(),endDateFormated.getMonth(),endDateFormated.getDate())
        const idFormated = employeeCode +  startDateFormated.getFullYear() + startDateFormated.getMonth() + startDateFormated.getDate() + id

        const newRow = [
          createdAt,
          idFormated,
          employeeCode,
          employeeName,
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
            entry[9]
          ];

          destinationSS.appendRow(row);
        }
      }
  }
}
