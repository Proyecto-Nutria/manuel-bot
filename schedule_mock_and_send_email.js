function schedule_mock_and_send_email(event) { // eslint-disable-line
  var activeSheet = event.source.getActiveSheet()

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

    if (currentTypeCell === typeHour) {
      var hour = getValueOf(activeSheet, currentRow, currentCol)
      var roomAssigned = getValueOf(activeSheet, roomRow, currentCol) !== empty
      var newInterview = currentTypeCell === typeHour &&
        currentCol < limitOfColumns &&
        hour !== '' &&
        !roomAssigned

      if (roomAssigned === true || newInterview === true) {
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
        } else if (newInterview) {
          var interviewer = getValueOf(activeSheet, interviewerRow, currentCol)
          if (isNumeric(interviewer)) {
            createAlert('Maybe you swap hours and your name')
          } else if (userEmail !== '') {
            createGoogleFolderFor(discordUser)
            var docName = createGoogleDocFor(discordUser)
            var docUrl = getGoogleDocBy(docName)
            var formattedRoom = 'room-' + findAvailableRoom(activeSheet, currentCol)
            var newEmailMessage = getBodyEmail(discordUser, day, hour, formattedRoom, docUrl, newInterviewEmailType)
            var newEmailSubject = 'Nutria Interview confirmation email'
            sendEmailTo(userEmail, newEmailSubject, newEmailMessage)
            var interviewerEmail = getEmailOf(event, interviewer, interviewersSheetName)
            createEventAndInvite(formattedRoom, day, hour, interviewerEmail)
            paintAndUpdateCells(activeSheet, currentCol, interviewerRow, googleDocUrlRow, roomRow, docUrl, formattedRoom)
          }
        }
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
    .getValues()[0][0]
    .replace(/(^\s+|\s+$)/g, ' ')
}

function getValueOf (activeSheet, row, col) {
  return activeSheet.getRange(row, col).getValue()
}

// Todo: Validate the number of rooms
function findAvailableRoom (activeSheet, columnDay) {
  const discordUserCol = 'P'
  var discordUserIndex = 4
  var indexRoom = 8
  var possibleRoom = 1

  while (true) {
    if (getInfoWithNoSpacesOF(activeSheet, discordUserCol, discordUserIndex) === '') break
    var room = getValueOf(activeSheet, indexRoom, columnDay)
    if (room !== '') possibleRoom += 1
    discordUserIndex += 5
    indexRoom += 5
  }

  return possibleRoom.toString()
}

function getUrlOfCell (activeSheet, row, col) {
  return /"(.*?)"/.exec(activeSheet.getRange(row, col).getFormulaR1C1())[1]
}

function createEventAndInvite (
  room,
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
    Logger.log('Error parsing the hour ' + interviewHourCleaned)
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

  CalendarApp.getCalendarById(
    mockCalendarId
  ).createEvent(
    'Mock Interview',
    new Date(
      today.getFullYear(),
      todayMonth,
      interviewDayCleaned,
      hour + 2,
      minutes,
      0
    ),
    new Date(
      today.getFullYear(),
      todayMonth,
      interviewDayCleaned,
      hour + 3,
      minutes,
      0
    ),
    {
      location: room,
      guests: interviewerEmail,
      sendInvites: true
    }
  )
}

function sendEmailTo (emailAddress, subject, message) {
  MailApp.sendEmail(emailAddress, subject, message)
}

function createAlert (message) {
  SpreadsheetApp.getUi().alert(message)
}

function createGoogleFolderFor (discordUser) {
  const feedbackFolderName = 'Feedback Docs'
  var folderDiscordUser = DriveApp.getFoldersByName(discordUser)
  if (!folderDiscordUser.hasNext()) {
    DriveApp.getFoldersByName(feedbackFolderName)
      .next()
      .createFolder(discordUser)
  }
}

function createGoogleDocFor (discordUser) {
  var docName = discordUser + '_' + (Math.random() * 50).toString()
  var docFile = DriveApp.getFileById(DocumentApp.create(docName).getId())
  docFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.EDIT
  )
  // Copy doc to the directory we want it to be in. Delete it from root.
  DriveApp.getFoldersByName(discordUser).next().addFile(docFile)
  DriveApp.getRootFolder().removeFile(docFile)

  return docName
}

function getGoogleDocBy (name) {
  return 'https://docs.google.com/document/d/' +
  DriveApp.getFilesByName(name).next().getId() +
  '/'
}

function paintAndUpdateCells (activeSheet, currentCol, interviewerRow, googleDocUrlRow, roomRow, docUrl, formattedRoom) {
  const numberOfElementsToPaint = 4
  activeSheet
    .getRange(interviewerRow, currentCol, numberOfElementsToPaint)
    .setBackgroundRGB(248, 255, 171)
  activeSheet
    .getRange(googleDocUrlRow, currentCol)
    .setFormula('=HYPERLINK("' + docUrl + '";"Docs Link")')
  activeSheet.getRange(roomRow, currentCol).setValue(formattedRoom)
}

function getEmailOf (event, id, sheetName) {
  const interviewersSheetName = 'Interviewers'
  const interviewersEmailRow = 'C'
  const intervieweeEmailRow = 'D'

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
      'B',
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
  if (type === 2) message = 'Your interview has been reschedule to '
  return 'Hi ' +
    discordUser.slice(0, -5) +
    ',\n\n' +
    message +
    day +
    ' at ' +
    hour +
    ' PT.\n\n' +
    'Place: ' +
    formattedRoom +
    '\n' +
    'Doc: ' +
    docUrl +
    '\n\n' +
    'Please confirm.\n\n' +
    'Best,\n' +
    'Your friends at Nutria'
}
