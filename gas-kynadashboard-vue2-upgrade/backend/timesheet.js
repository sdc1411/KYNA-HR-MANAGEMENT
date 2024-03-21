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
  const rowData = [];

  try {
    const conn = Jdbc.getConnection(dbUrl, username, password);
    const stmt = conn.createStatement();
    const results = stmt.executeQuery(queryString);

    while (results.next()) {
      const employeeCode = results.getString('employeeCode');
      const startDate = new Date(results.getString('startDate'));
      const endDate = new Date(results.getString('endDate'));
      const startShift = new Date(results.getString('startShift'));
      const endShift = new Date(results.getString('endShift'));
      const startShiftFormat = Utilities.formatDate(startShift, Session.getScriptTimeZone(), 'h:mm:ss a')
      const endShiftFormat = Utilities.formatDate(endShift, Session.getScriptTimeZone(), 'h:mm:ss a')

      rowData.push([employeeCode, startDate, endDate, startShiftFormat, endShiftFormat, 'Chấm công vân tay', 'Approved']);
    }

    results.close();
    stmt.close();
    conn.close();
  } catch (e) {
    console.error('Error: ' + e);
  }

  const ss = SpreadsheetApp.openById(CONFIG.DATABASE.TIME_SHEET)
  const destinationSS = ss.getSheetByName('timeSheet');

  if (rowData.length > 0) {
    const destinationValue = destinationSS.getDataRange()
    const lastRow = destinationValue.getLastRow()
    const lastCol = destinationValue.getLastColumn()
    destinationSS.getRange(lastRow + 1, 1, rowData.length, lastCol).setValues(rowData)
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






function getEvents(date, employeeCode) {
  date = '2023-09-26'
  employeeCode = 'K-0001'
  const resources = [
        { employeeCode: 'K-0001',employeeName: 'John', events: [{ date: '2023-09-26', endDate: '2023-09-26', time: '08:00', duration: 200, endShift: '17:30', title: 'working', bgcolor: 'green' },{ date: '2023-09-26', endDate: '2023-09-26', time: '08:00', duration: 150, endShift: '17:30', title: 'nghỉ phép', bgcolor: 'orange' }]},
        { employeeCode: 'K-0002',employeeName: 'Linda', events: [{ date: '2023-09-26', endDate: '2023-09-30', time: '08:00', endShift: '17:30', day: 3, title: 'nghỉ phép', bgcolor: 'orange'}]},
        { employeeCode: 'K-0003',employeeName: 'Mary', events: [{ date: '2023-09-29', endDate: '2023-09-29', time: '08:00', endShift: '17:30', title: 'working', bgcolor: 'green' }] },
        { employeeCode: 'K-0004',employeeName: 'Susan', events: [{ date: '2023-09-30', endDate: '2023-09-30', time: '12:00', endShift: '22:00', title: 'working', bgcolor: 'green' }] },
        { employeeCode: 'K-0005',employeeName: 'Olivia' }
      ]
  const currentDate = date
  const events = []
  resources.forEach((resource) => {
    if (resource.events && resource.employeeCode === employeeCode) {
      for (let i = 0; i < resource.events.length; ++i) {
        let added = false
        if (resource.events[i].date === date && resource.events[i].date === resource.events[i].endDate) {
          if (resource.events[i].time) {
            const startTimeTime = new Date(`${resource.events[i].date} ${resource.events[i].time}`);
            const endTimeTime = new Date(`${resource.events[i].endDate} ${resource.events[i].endShift}`);
            const startTime = startTimeTime.getTime()
            const endTime = endTimeTime.getTime()

            for (let j = 0; j < events.length; ++j) {
              if (events[j].time) {
                const startTimeTime2 = new Date(`${resource.events[j].date} ${resource.events[j].time}`);
                const endTimeTime2 = new Date(`${resource.events[j].endDate} ${resource.events[j].endShift}`);
                const startTime2 = startTimeTime2.getTime();
                const endTime2 = endTimeTime2.getTime();

                if ((startTime >= startTime2 && startTime <= endTime2) || (endTime >= startTime2 && endTime <= endTime2)) {
                  // Determine 'left' or 'right' based on existing events
                  if (startTime < startTime2) {
                    events[j].side = 'left';
                    resource.events[i].side = 'right';
                  } else {
                    events[j].side = 'right';
                    resource.events[i].side = 'left';
                  }
                  events.push(resource.events[i])
                  added = true;
                  break;
                }
              }
            }
            // If no overlap was found, add the event without setting 'side'
            if (!added) {
              resource.events[i].side = undefined
              events.push(resource.events[i])
            }
          }
        }
        else if (resource.events[i].date === date && resource.events[i].endDate > resource.events[i].date) {
          // check for overlapping dates
          const date = resource.events[i].date
          const endDate = resource.events[i].endDate
          if (currentDate >= date && currentDate <= endDate) {
            events.push(resource.events[i])
            added = true
          }
        }
      }

    }
  })
  console.log(events)
  return events
}


function getEvents2(date, employeeCode) {
  date = '2023-09-26'
  employeeCode = 'K-0001'
  const resources = [
    { employeeCode: 'K-0001', employeeName: 'John', events: [{ startDate: '2023-09-26', endDate: '2023-09-26', startShift: '08:00', endShift: '17:30', title: 'working', bgcolor: 'green' }, { startDate: '2023-09-26', endDate: '2023-09-26', startShift: '08:00', endShift: '17:30', title: 'nghỉ phép', bgcolor: 'orange' }, { startDate: '2023-09-26', endDate: '2023-09-29', startShift: '08:00', endShift: '17:30', title: 'nghỉ ốm', bgcolor: 'lewi' }] },
    { employeeCode: 'K-0002', employeeName: 'Linda', events: [{ startDate: '2023-09-26', endDate: '2023-09-30', startShift: '08:00', endShift: '17:30', title: 'nghỉ phép', bgcolor: 'orange' }] },
    { employeeCode: 'K-0003', employeeName: 'Mary', events: [{ startDate: '2023-09-29', endDate: '2023-09-29', startShift: '08:00', endShift: '17:30', title: 'working', bgcolor: 'green' }] },
    { employeeCode: 'K-0004', employeeName: 'Susan', events: [{ startDate: '2023-09-30', endDate: '2023-09-30', startShift: '12:00', endShift: '22:00', title: 'working', bgcolor: 'green' }] },
    { employeeCode: 'K-0005', employeeName: 'Olivia' }
  ]
  const events = [];
  resources.forEach((resource) => {
    if (resource.events && resource.employeeCode === employeeCode) {
      resource.events.forEach((event) => {
        if (event.startDate === date) {
          const formattedEvent = {
            date: event.startDate,
            title: event.title,
            time: event.startShift,
            bgcolor: event.bgcolor,
          };
          events.push(formattedEvent);
        }
      });
    }
  });
  console.log(events)
  return events;
}


