/**
 * フォーム送信時に実行される
 * @param e Form submit event
 */
function onFormSubmit(e) {
  const response: GoogleAppsScript.Forms.FormResponse = e.response
  const email = response.getRespondentEmail()
  const date = getDate()
  const calendarEvent = searchCalendarEvent(date)
  calendarEvent.addGuest(email)
}

/**
 * フォーム送信時に実行される
 * @param e Spreadsheet Form submit event
 */
function onFormSubmitToSpreadSheet(e) {
  const range: GoogleAppsScript.Spreadsheet.Range = e.range
  const values: string[] = e.values
  const mailColumn = getMailColumn(range.getSheet())
  const email = values[mailColumn]
  const date = getDate()
  const calendarEvent = searchCalendarEvent(date)
  calendarEvent.addGuest(email)
}

/**
 * スプレッドシートのメールアドレス列番号取得
 * @param sheet Spreadsheet
 */
function getMailColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn())
  const values = firstRow.getValues()[0]
  for (let i = 0; i < values.length; i++) {
    if (values[i].toString().indexOf('メールアドレス') !== -1) {
      return i
    }
  }
}

/**
 * プロパティに設定された日付を取得する
 * @param sheet Spreadsheet
 */
function getDate(): Date {
  const property = PropertiesService.getScriptProperties()
  return new Date(property.getProperty('EVENT_DATE'))
}

/**
 * 入力された日付の該当イベントを取得
 * @param eventDay イベントの日付
 */
function searchCalendarEvent(eventDay: Date): GoogleAppsScript.Calendar.CalendarEvent {
  const property = PropertiesService.getScriptProperties()
  const calendar = CalendarApp.getCalendarById(property.getProperty('CALENDAR_ID'))
  const events = calendar.getEventsForDay(eventDay)
  for (let event of events) {
    if (event.getTitle().indexOf(property.getProperty('EVENT_NAME')) !== -1) {
      return event
    }
  }
}

/**
 * Form にトリガーを作成する
 */
function createFormTrigger() {
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create()
}

/**
 * Spreadsheet にトリガーを作成する
 */
function createSpreadsheetTrigger() {
  ScriptApp.newTrigger('onFormSubmitToSpreadSheet')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create()
}