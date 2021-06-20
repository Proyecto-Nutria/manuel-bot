function schedule_mock_and_send_email (event) { // eslint-disable-line
  var activeSpreadsheet = event.source
  var activeSheet = activeSpreadsheet.getActiveSheet()

  if (activeSheet.getName() === 'Schedule') {
    const empty = ''
    const dayRow = 3
    const typeHour = 2
    const limitOfColumns = 16
    const interviewersSheetName = 'Interviewers'
    const logSheetName = 'Log'
    const newInterviewEmailType = 1
    const updateInterviewEmailType = 2

    var range = event.range
    var currentRow = parseInt(range.getRow())
    var currentCol = parseInt(range.getColumn())
    var currentTypeCell = (currentRow - 4) % 5

    var roomRow = currentRow + 2
    var googleDocUrlRow = currentRow + 1
    var interviewerRow = currentRow - 1

    if (currentTypeCell === typeHour && currentCol < limitOfColumns) {
      var hour = getValueOf(activeSheet, currentRow, currentCol)
      var roomAssigned = getValueOf(activeSheet, roomRow, currentCol) !== empty
      var newInterview = currentTypeCell === typeHour &&
        hour !== '' &&
        !roomAssigned

      if (roomAssigned === true || newInterview === true) {
        // "Sin Mensaje" - color
        var colorAccent1 = activeSpreadsheet.getSpreadsheetTheme().getConcreteColor(
          SpreadsheetApp.ThemeColorType.ACCENT1 // eslint-disable-line
        )
        // "Mensaje enviado" - color
        var colorAccent3 = activeSpreadsheet.getSpreadsheetTheme().getConcreteColor(
          SpreadsheetApp.ThemeColorType.ACCENT3 // eslint-disable-line
        )

        paintCells(activeSheet, interviewerRow, currentCol, colorAccent1)

        var discordUser = getDiscordUserOf(activeSheet, currentRow)
        var userEmail = getEmailOf(event, discordUser, logSheetName)
        var day = getValueOf(activeSheet, dayRow, currentCol)

        // Interview Updated
        if (roomAssigned) {
          var prevGoogleDocUrl = getUrlOfCell(activeSheet, googleDocUrlRow, currentCol)
          var room = getValueOf(activeSheet, roomRow, currentCol)
          var updateEmailMessage = getBodyEmail(discordUser, day, hour, room, prevGoogleDocUrl, updateInterviewEmailType)
          var subject = 'Updates for your ' + day + ' mock interview'
          sendEmailTo(userEmail, subject, updateEmailMessage)
          // Todo: Update calendar event
        } else if (newInterview) {
          var interviewer = getValueOf(activeSheet, interviewerRow, currentCol)
          if (isNumeric(interviewer)) {
            createAlert('Maybe you swap hours and your name')
          } else if (userEmail !== '') {
            createGoogleFolderFor(discordUser)
            var docName = createGoogleDocFor(discordUser)
            var docUrl = getGoogleDocBy(docName)
            var formattedRoom = 'room-' + findAvailableRoom(activeSheet, currentCol, getValueOf(activeSheet, currentRow, currentCol))
            var newEmailMessage = getBodyEmail(discordUser, day, hour, formattedRoom, docUrl, newInterviewEmailType)
            var newEmailSubject = 'Nutria Interview confirmation email'
            sendEmailTo(userEmail, newEmailSubject, newEmailMessage)
            var interviewerEmail = getEmailOf(event, interviewer, interviewersSheetName)
            createEventAndInvite(discordUser, formattedRoom, docUrl, day, hour, interviewerEmail)
            updateInterviewInfo(activeSheet, currentCol, googleDocUrlRow, roomRow, docUrl, formattedRoom)
            writeLogEntry(event, discordUser, day, docUrl, interviewer, logSheetName)
          }
        }

        paintCells(activeSheet, interviewerRow, currentCol, colorAccent3)
      }
    }
  }
}

function isNumeric (str) {
  return /\d/.test(str)
}

function getDiscordUserOf (activeSheet, currentRow) {
  return getInfoWithNoSpacesOF(activeSheet, 'P', currentRow - 2)
}

function getInfoWithNoSpacesOF (sheet, letterCell, numberCell) {
  return sheet
    .getRange(letterCell + numberCell.toString())
    .getValues()[0][0].toString()
    .replace(/(^\s+|\s+$)/g, ' ')
}

function intervalsIntersect (interval1, interval2) {
  if (interval1[1] <= interval2[0] || interval2[1] <= interval1[0]) {
    return false
  } else {
    return true
  }
}

function findSpaceInRoom (busyTimeIntervalsForRoom, newInterviewInterval) {
  for (var i = 0; i < busyTimeIntervalsForRoom.length; i++) {
    if (intervalsIntersect(busyTimeIntervalsForRoom[i], newInterviewInterval)) return false
  }
  return true
}

function toMinutesInADay (timeAsString) {
  timeAsString = String(timeAsString).toLowerCase()
  var parsedTime = timeAsString.match(/\d{1,2}(:\d{2})?/)
  if (!parsedTime) {
    Logger.log('Wrong time format') // eslint-disable-line
    return -1
  }
  parsedTime = parsedTime[0].split(':')
  var hours = parseInt(parsedTime[0])
  var minutes = parsedTime.length < 2 ? 0 : parseInt(parsedTime[1])

  if (hours === 12 && timeAsString.includes('am')) {
    hours = 0
  } else if (hours < 12 && timeAsString.includes('pm')) {
    hours += 12
  }
  return hours * 60 + minutes
}

function getRoomId (rawText) {
  return parseInt(rawText.match(/\d+/)[0])
}

function findAvailableRoom (activeSheet, columnDay, newTime) {
  const discordUserCol = 'P'
  var discordUserIndex = 4
  var indexRoom = 8
  var roomsLimit = 10
  var rooms = []
  var auxcnt = 0

  var interviewDurationInMinutes = 75
  for (var i = 0; i < roomsLimit; i++) {
    rooms.push([[0, 0]])
  }

  var maxLimit = 100
  var it = 0
  while (it < maxLimit) {
    if (getInfoWithNoSpacesOF(activeSheet, discordUserCol, discordUserIndex) === '') break
    var room = getValueOf(activeSheet, indexRoom, columnDay)
    if (room !== '') {
      var roomId = getRoomId(room) - 1
      var time = toMinutesInADay(getValueOf(activeSheet, indexRoom - 2, columnDay))
      rooms[roomId].push([time, time + interviewDurationInMinutes])
      auxcnt++
    }
    discordUserIndex += 5
    indexRoom += 5
    it++
  }

  var availableRoom = null
  newTime = toMinutesInADay(newTime)
  var newTimeInterval = [newTime, newTime + interviewDurationInMinutes]
  for (var limit = 0; limit < roomsLimit; limit++) {
    if (findSpaceInRoom(rooms[limit], newTimeInterval)) {
      availableRoom = limit + 1
      break
    }
  }

  if (availableRoom) {
    return availableRoom.toString()
  } else {
    return (auxcnt + 1).toString()
  }
}

function getUrlOfCell (activeSheet, row, col) {
  return /"(.*?)"/.exec(activeSheet.getRange(row, col).getFormulaR1C1())[1]
}

function createEventAndInvite (
  discordUser,
  room,
  docURL,
  interviewDay,
  interviewHour,
  interviewerEmail
) {
  var hour = 0
  var minutes = 0
  var interviewHourCleaned = interviewHour.trim().toLowerCase()
  var splittedHour = interviewHourCleaned.split(':')
  var am = false
  var meridiem = ''

  // 9pm
  if (splittedHour.length === 1) {
    var hourWithMeridiem = interviewHourCleaned.split(/(\d+)/g)
    hour = parseInt(hourWithMeridiem[1])
    meridiem = hourWithMeridiem[2]
  } else {
    hour = parseInt(splittedHour[0])
    var minutesWithMeridiem = splittedHour[1].split(/(\d+)/g)
    minutes = parseInt(minutesWithMeridiem[1])
    meridiem = minutesWithMeridiem[2]
  }

  if (typeof meridiem === 'undefined') {
    Logger.log('Error parsing the hour ' + interviewHourCleaned) // eslint-disable-line
    return
  }

  if (meridiem.includes('am') === true) am = true
  if (am === false && hour < 12) hour = hour + 12
  if (am === true && hour === 12) hour = hour - 12

  var mockCalendarId = '6hf4jhvprg5d8b66lu4ld93hho@group.calendar.google.com'
  var today = new Date()
  var todayDay = today.getDate()
  var todayMonth = today.getMonth()
  var interviewDayCleaned = interviewDay.replace(/\D/g, '')

  if (interviewDayCleaned < todayDay) todayMonth += 1

  const localTime = new Date(
    today.getFullYear(),
    todayMonth,
    interviewDayCleaned,
    hour,
    minutes,
    0
  )
  const ptTime = new Date(localTime.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }))
  const diffWithPT = localTime.getTime() - ptTime.getTime()

  CalendarApp.getCalendarById( // eslint-disable-line
    mockCalendarId
  ).createEvent(
    'Mock Interview: ' + discordUser,
    new Date(localTime.getTime() + diffWithPT),
    new Date(localTime.getTime() + diffWithPT + 3600000),
    {
      description: '<a href="' + docURL + '">Docs link</a>',
      location: room,
      guests: interviewerEmail,
      sendInvites: true
    }
  )
}

function sendEmailTo (emailAddress, subject, message) {
  MailApp.sendEmail(emailAddress, subject, message) // eslint-disable-line
}

function createAlert (message) {
  SpreadsheetApp.getUi().alert(message) // eslint-disable-line
}

function createGoogleFolderFor (discordUser) {
  const feedbackFolderName = 'Feedback Docs'
  var folderDiscordUser = DriveApp.getFoldersByName(discordUser) // eslint-disable-line
  if (!folderDiscordUser.hasNext()) {
    DriveApp.getFoldersByName(feedbackFolderName) // eslint-disable-line
      .next()
      .createFolder(discordUser)
  }
}

function createGoogleDocFor (discordUser) {
  var docName = discordUser + '_' + (Math.random() * 50).toString()
  var docFile = DriveApp.getFileById(DocumentApp.create(docName).getId()) // eslint-disable-line
  docFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK, // eslint-disable-line
    DriveApp.Permission.EDIT // eslint-disable-line
  )
  // Copy doc to the directory we want it to be in. Delete it from root.
  DriveApp.getFoldersByName(discordUser).next().addFile(docFile) // eslint-disable-line
  DriveApp.getRootFolder().removeFile(docFile) // eslint-disable-line

  return docName
}

function getGoogleDocBy (name) {
  return 'https://docs.google.com/document/d/' +
  DriveApp.getFilesByName(name).next().getId() + // eslint-disable-line
  '/'
}

function paintCells (activeSheet, interviewerRow, currentCol, color) {
  const numberOfElementsToPaint = 4
  activeSheet
    .getRange(interviewerRow, currentCol, numberOfElementsToPaint)
    .setBackgroundObject(color)
}

function updateInterviewInfo (activeSheet, currentCol, googleDocUrlRow, roomRow, docUrl, formattedRoom) {
  activeSheet
    .getRange(googleDocUrlRow, currentCol)
    .setFormula('=HYPERLINK("' + docUrl + '";"Docs Link")')
  activeSheet.getRange(roomRow, currentCol).setValue(formattedRoom)
}

function getEmailOf (event, id, sheetName) {
  const interviewersSheetName = 'Interviewers'
  const interviewersEmailRow = 'C'
  const intervieweeEmailRow = 'D'
  const idColumn = 'B'

  var userEmail = ''
  var currentEntry = 3
  var emailRow = intervieweeEmailRow
  var sheet = event.source.getSheetByName(sheetName)

  if (sheetName === interviewersSheetName) {
    currentEntry = 2
    emailRow = interviewersEmailRow
  }

  while (true) {
    var currentEntryAsStr = currentEntry.toString()
    var user = getInfoWithNoSpacesOF(
      sheet,
      idColumn,
      currentEntryAsStr
    )
    if (user === '') {
      createAlert('Email of ' + id + ' not found in ' + sheetName)
      break
    } else if (user === id) {
      userEmail = getInfoWithNoSpacesOF(
        sheet,
        emailRow,
        currentEntryAsStr
      )
      break
    }
    currentEntry += 1
  }
  return userEmail
}

function getBodyEmail (discordUser, day, hour, formattedRoom, docUrl, type) {
  var message = 'We have scheduled an interview for '
  if (type === 2) message = 'Your interview has been rescheduled to '
  return 'Hi ' +
    discordUser +
    '\n\n' +
    message +
    day +
    ' , ' +
    hour +
    ' PT.\n\n' +
    'Place: ' +
    formattedRoom +
    '\n' +
    'Doc: ' +
    docUrl +
    '\n\n' +
    'Please confirm by replying to this email (You need to specifically write the word "confirm" in your reply).\n\n' +
    'Best,\n' +
    'Your friends at Nutria'
}

function getValueOf (activeSheet, row, col) {
  return activeSheet.getRange(row, col).getValue()
}

function getCompressedDate (interviewDay) {
  var today = new Date()
  var day = today.getDate()
  var month = today.getMonth() + 1
  var year = today.getFullYear().toString().substr(2, 2)

  var interviewDayCleaned = interviewDay.replace(/\D/g, '')
  if (interviewDayCleaned.length < 2) interviewDayCleaned = '0' + interviewDayCleaned

  if (interviewDayCleaned < day) month += 1

  month = month.toString()
  if (month.length < 2) month = '0' + month

  return `${interviewDayCleaned}/${month}/${year}`
}

function columnToLetter (column) {
  var temp = ''
  var letter = ''
  while (column > 0) {
    temp = (column - 1) % 26
    letter = String.fromCharCode(temp + 65) + letter
    column = (column - temp - 1) / 26
  }
  return letter
}

function writeLogEntry (event, discordUser, interviewDay, docUrl, interviewer, logSheetName) {
  var currentEntry = 3
  var logSheet = event.source.getSheetByName(logSheetName)
  var completeDate = getCompressedDate(interviewDay)
  while (true) {
    var currentEntryStr = currentEntry.toString()
    var logDiscordUser = getInfoWithNoSpacesOF(logSheet, 'B', currentEntryStr)
    if (logDiscordUser === '') {
      Logger.log('User not found at Log Sheets') // eslint-disable-line
      createAlert('Email of ' + discordUser + ' not found in Log')
      break
    } else if (logDiscordUser === discordUser) {
      var currentColumn = 6
      while (true) {
        const cellDate = getInfoWithNoSpacesOF(logSheet, columnToLetter(currentColumn), currentEntryStr)
        if (cellDate === '') break
        else currentColumn += 6
      }
      logSheet.getRange(currentEntry, currentColumn).setValue(completeDate)
      logSheet.getRange(currentEntry, currentColumn + 4).setFormula('=HYPERLINK("' + docUrl + '";"Link")')
      logSheet.getRange(currentEntry, currentColumn + 5).setValue(interviewer)
      break
    }
    currentEntry += 1
  }
}

module.exports = {
  isNumeric,
  getRoomId,
  columnToLetter,
  getInfoWithNoSpacesOF,
  intervalsIntersect,
  findSpaceInRoom,
  toMinutesInADay,
  getUrlOfCell
}
