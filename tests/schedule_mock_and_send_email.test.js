const { isNumeric,
  getRoomId,
  columnToLetter,
  getInfoWithNoSpacesOF,
  intervalsIntersect,
  findSpaceInRoom,
  toMinutesInADay,
  getUrlOfCell,
  getCompressedDate
} = require('../schedule_mock_and_send_email')

test('Is Numeric', () => {
  expect(isNumeric('1')).toBe(true)
  expect(isNumeric('99')).toBe(true)
  expect(isNumeric('A')).toBe(false)
  expect(isNumeric('A B')).toBe(false)
  expect(isNumeric(' ')).toBe(false)
  expect(isNumeric('undefined')).toBe(false)
})

test('Get Room Id', () => {
  expect(getRoomId('room-10')).toBe(10)
  expect(getRoomId('room-1')).toBe(1)
})

test('Column To Letter', () => {
  expect(columnToLetter(1)).toBe('A')
  expect(columnToLetter(26)).toBe('Z')
  expect(columnToLetter(27)).toBe('AA')
  expect(columnToLetter(52)).toBe('AZ')
})

test('Get Info With no Spaces', () => {
  var sheet = {
    getRange: function () {
      return { getValues: function () { return [['Some Val']] } }
    }
  }
  var emptySheet = {
    getRange: function () {
      return { getValues: function () { return [['']] } }
    }
  }

  expect(getInfoWithNoSpacesOF(sheet, 'B', 0)).toBe('Some Val')
  expect(getInfoWithNoSpacesOF(emptySheet, 'B', 0)).toBe('')
})

test('Test Intervals', () => {
  expect(intervalsIntersect([1, 3], [2, 4])).toBe(true)
  expect(intervalsIntersect([1, 3], [4, 5])).toBe(false)
})

test('Find Space In Room', () => {
  expect(findSpaceInRoom([[1, 2], [3, 6], [9, 14]], [6, 9])).toBe(true)
  expect(findSpaceInRoom([[1, 5], [5, 9]], [2, 3])).toBe(false)
})

test('Day To Minutes', () => {
  expect(toMinutesInADay('12:00 am')).toBe(0)
  expect(toMinutesInADay('12:00 AM')).toBe(0)
  expect(toMinutesInADay('12:00 aM')).toBe(0)
  expect(toMinutesInADay('12:00am')).toBe(0)
  expect(toMinutesInADay('12:00AM')).toBe(0)
  expect(toMinutesInADay('12:00aM')).toBe(0)

  expect(toMinutesInADay('12:00 pm')).toBe(720)
  expect(toMinutesInADay('12:00 PM')).toBe(720)
  expect(toMinutesInADay('12:00 pM')).toBe(720)
  expect(toMinutesInADay('12:00pm')).toBe(720)
  expect(toMinutesInADay('12:00PM')).toBe(720)
  expect(toMinutesInADay('12:00pM')).toBe(720)

  expect(toMinutesInADay('10:30 AM')).toBe(630)
  expect(toMinutesInADay('10:30am')).toBe(630)
  expect(toMinutesInADay('1:30am')).toBe(90)

  expect(toMinutesInADay('10:30 PM')).toBe(1350)
  expect(toMinutesInADay('10:30pm')).toBe(1350)
  expect(toMinutesInADay('1:30pm')).toBe(810)
})

test('Get Url Of Cell', () => {
  const googleDocUrl = 'https://docs.google.com/document/d/000/'
  var activeSheet = {
    getRange: function () {
      return { getFormulaR1C1: function () { return '=HYPERLINK("' + googleDocUrl + '";"Docs Link")' } }
    }
  }
  expect(getUrlOfCell(activeSheet, 0, 0)).toBe(googleDocUrl)
})
