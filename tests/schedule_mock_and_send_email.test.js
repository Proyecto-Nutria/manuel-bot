const { isNumeric, getRoomId, columnToLetter, getInfoWithNoSpacesOF } = require('../schedule_mock_and_send_email')

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