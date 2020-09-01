function checkGmailUsingChron () { // eslint-disable-line
  updateSheetIfConfirmationEmailIn(getUnReadEmails())
}

function getUnReadEmails () {
    return GmailApp.search('is:unread') // eslint-disable-line
}

function updateSheetIfConfirmationEmailIn (unreadEmails) {
  const nutriaEmail = 'proyecto.nutria.escom@gmail.com'
  const nutriaNewSubject = 'Nutria Interview'
  const nutriaUpdateSubject = 'Updates'

  for (var unreadIndex = 0; unreadIndex < unreadEmails.length; unreadIndex++) {
    var currentEmail = unreadEmails[unreadIndex]
    var emailSubject = currentEmail.getFirstMessageSubject()

    if (emailSubject.includes(nutriaNewSubject) || emailSubject.includes(nutriaUpdateSubject)) {
      var discordUser = ''
      var interviewDay = ''
      var nutriaMessage = false
      var confirmation = false
      var allMessages = currentEmail.getMessages()
      var lastMessage = allMessages[allMessages.length - 1]
      var bodyOfTheEmailSplittedByLine = lastMessage.getPlainBody().split('\r\n')

      for (var lineIndex = 0; lineIndex < bodyOfTheEmailSplittedByLine.length; lineIndex++) {
        var currentLine = bodyOfTheEmailSplittedByLine[lineIndex]
        if (currentLine !== '') {
          if (nutriaMessage === false) {
            if (currentLine.includes(nutriaEmail)) {
              nutriaMessage = true
            } else if (currentLine.toLowerCase().includes('confirm')) {
              confirmation = true
            }
          }

          if (nutriaMessage === true && confirmation === true) {
            if (interviewDay !== '') break
            if (currentLine.toLowerCase().includes('hi')) discordUser = currentLine.split(' ')[2]

            if (currentLine.toLowerCase().includes('scheduled')) {
              interviewDay = currentLine.split('for')[1].split('at')[0]
            }

            if (currentLine.toLowerCase().includes('reschedule')) {
              interviewDay = currentLine.split('to')[1].split('at')[0]
            }
          }
        }
      }

      if (interviewDay !== '' && discordUser !== '') {
        var sheetSearch = DriveApp.getFilesByName('Schedule for Mock Interviews') // eslint-disable-line
        currentEmail.markRead()

        if (sheetSearch.hasNext()) {
          var scheduleSheet = SpreadsheetApp.open(sheetSearch.next()).getSheets()[0] // eslint-disable-line

          for (var column = 2; column < 16; column++) {
            var dayOfSchedule = scheduleSheet.getRange(3, column).getValue()

            if (interviewDay.includes(dayOfSchedule)) {
              var rowOfDiscordUser = findRowOfDiscordUser(scheduleSheet, discordUser)

              if (rowOfDiscordUser !== -1) {
                const numberOfElementsToPaint = 4
                scheduleSheet
                  .getRange(rowOfDiscordUser + 1, column, numberOfElementsToPaint)
                  .setBackgroundRGB('100', '221', '23')
              }
            }
          }
        }
      }
    }
  }
}

function findRowOfDiscordUser (scheduleSheet, discordUser) {
  const discordUserColumn = 16
  var currentRow = 4
  while (true) {
    var currentUser = scheduleSheet.getRange(currentRow, discordUserColumn).getValue()
    if (currentUser === '') return -1
    if (currentUser === discordUser) return currentRow
    currentRow += 5
  }
}
