// function createTimeDrivenTriggers()
// {
//   ScriptApp.newTrigger('checkDeadLinesDverdue')
//       .timeBased()
//       .atHour(9)
//       .create();
// }


const SHEETS_TITLES =
{
  DEMAND:     'Demand Backlog',
  SUPPLY:     'Supply Backlog',
  BALANCE:    'Balance Backlog',
  EXECUTION:  'NONE',
  ANALYTICKS: 'Analytics backlog',
  PRODUCT:    'Product backlog',
  PPL_OPS:    'Common ppl ops',
  PICKERS:    'Pickers -->',
};

const INITIAL_COLUMNS_TITLES = 
{
  TASK_NAME:        'Проект / задача',  // 1
  TASK_STREAM:      'Стрим',            // 2
  TASK_DESCRIPTION: 'Описание',         // 3
  TASK_OWNER:       'Ответственный',    // 4
  TASK_STATUS:      'Статус',           // 5
  TASK_DEADLINE:    'ETA',              // 6
  TASK_COMMENTS:    'Комменты',         // 7
};

const OUTPUT_COLUMNS_TITLES = 
{
  TASK_NAME:        'Проект / задача',  // 1
  TASK_INITIATOR:   'Кто инициатор',    // 2
  TASK_DESCRIPTION: 'Описание',         // 3
  TASK_STATUS:      'Статус',           // 4
};

function checkDeadLinesDverdue()
{
  // TODO
}

function convertInitiatorFromSheetName(sheet)
{
  console.log("convertInitiatorFromSheetName(sheet):", new Date());

  if (!sheet)
  {
    console.log("ERROR: convertInitiatorFromSheetName: !sheet", sheet);
    return null;
  }

  switch(sheet.getName())
  {
    case SHEETS_TITLES.DEMAND:    return 'Деманд';
    case SHEETS_TITLES.SUPPLY:    return 'Сапплай';
    case SHEETS_TITLES.BALANCE:   return 'Баланс';
    case SHEETS_TITLES.EXECUTION: return 'Сборка и выполнение';
    case SHEETS_TITLES.PPL_OPS:   return 'Внутренние операции';
    default:                      return sheet.getName();
  }
}

function getColumnNumberByName(sheet, name, rowIdx = 2)
{
  console.log("getColumnNumberByName(sheet, name):", new Date());

  if (!sheet || !name)
  {
    console.log("ERROR: getColumnNumberByName: !sheet || !name", sheet, name);
    return 0;
  }

  for (var i = 1; i < sheet.getMaxColumns(); ++i)
  {
    var currentValue = sheet.getRange(rowIdx, i).getValue().toString();

    if (currentValue)
    {
      if (currentValue.trim() == name.trim())
        return i;
    }
  }

  return 0;
}

function getLastFreeRow(sheet)
{
  console.log("getLastFreeRow(sheet):", new Date());

  if (!sheet)
  {
    console.log("ERROR: getLastFreeRow: !sheet", sheet);
    return null;
  }

  for (var i = 1; i < sheet.getMaxRows()-1; ++i)
  {
    var currentValue = sheet.getRange(i, 1).getValue().toString();
    var nextValue = sheet.getRange(i+1, 1).getValue().toString();

    if (currentValue.length === 0 && nextValue.length === 0)
        return i;
  }

  return 0;
}

function getSheetByName(allSheets, name)
{
  console.log("getSheetByName(allSheets, name):", new Date());

  if (!allSheets)
  {
    console.log("ERROR: getSheetByName: !allSheets", allSheets, name);
    return null;
  }

  for (var i = 0; i < allSheets.length; ++i)
  {
    if (allSheets[i].getName() == name)
      return allSheets[i];
  }

  return null;
}

function getCurDate()
{
  return new Date().toISOString().substring(0,10);
}

function copyCellData(fromSheet, toSheet, _fromRow, _fromColumnName, _toRow, _toColumnName, customText = null)
{
  console.log("copyCellData(fromSheet, toSheet, _fromRow, _fromColumnName, _toRow, _toColumnName, customText):", 
              new Date());

  if (!fromSheet || !toSheet)
  {
    console.log("ERROR: copyCellData: !fromSheet || !toSheet", fromSheet, toSheet)
    return false;
  }

  var fromColumn = getColumnNumberByName(fromSheet, _fromColumnName);
  var toColumn = getColumnNumberByName(toSheet, _toColumnName);

  if ((!fromColumn || !toColumn) && !customText)
  {
    console.log("ERROR: copyCellData: !fromColumn || !toColumn", fromColumn, toColumn)
    return false;
  }

  var newValue = customText ? customText.toString() : fromSheet.getRange(_fromRow, fromColumn).getValue();

  toSheet.getRange(_toRow, toColumn).setValue(newValue); 

  return true;
}

function findSameProjectInOtherSheet(sheet, prjName)
{
  console.log("findSameProjectInOtherSheet(sheet, prjName):", new Date());

  if (!sheet || !prjName)
  {
    console.log("ERROR: getColumnNumberByName: !sheet || !prjName", sheet, prjName);
    return 0;
  }

  var rows = getLastFreeRow(sheet);

  for (var i = 1; i < rows/*sheet.getMaxRows()-1*/; ++i)
  {
    if (sheet.getRange(i, 1).getValue().toString() == prjName)
      return true;
  }

  return false;
}

function getMailAdresses(sheet, mailTo)
{
  console.log("getMailAdresses(sheet):", new Date());

  if (!sheet)
  {
    console.log("ERROR: getMailAdresses: !sheet", sheet);
    return null;
  }

  var analitycsEmailsCol = getColumnNumberByName(sheet, 'ANALYTICS emails', 3);
  var productsEmailsCol = getColumnNumberByName(sheet, 'PRODUCT emails', 3);

  if (mailTo == 'analitycs')
  {
    if (!analitycsEmailsCol)
    {
      console.log("ERROR: getMailAdresses: !analitycsEmailsCol", analitycsEmailsCol);
      return null;
    }
  }

  if (mailTo == 'product')
  {
    if (!productsEmailsCol)
    {
      console.log("ERROR: getMailAdresses: !analitycsEmailsCol", analitycsEmailsCol);
      return null;
    }
  }

  var emailsLst = [];
  var column = mailTo == 'analitycs' ? analitycsEmailsCol : productsEmailsCol;

  for (var i = 4; i < sheet.getMaxRows(); ++i)
  {
    var value = sheet.getRange(i, column).getValue();
    var posibility = sheet.getRange(i, (column + 1)).getValue();

    if (!(value.length === 0) && posibility == 'YES')
      emailsLst.push(value);
  }

  console.log(emailsLst);
  return emailsLst.join(",");
}

// function onEdit() 
function onEdit(e) 
{
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = activeSpreadsheet.getSheets();
  var activeSheet = activeSpreadsheet.getActiveSheet();

  var cell = activeSheet.getCurrentCell();
  var cellValue = cell.getValue();

  if (cell.getColumn() == getColumnNumberByName(activeSheet, INITIAL_COLUMNS_TITLES.TASK_NAME))
  {
    if (cellValue && cellValue != "" && !(cellValue.includes('§')))
      cell.setValue("§ " + getCurDate() + '\n' + cellValue);
  }

  if (cell.getColumn() == getColumnNumberByName(activeSheet, INITIAL_COLUMNS_TITLES.TASK_STATUS))
  {
    if (cellValue == 'Analytics needed')
    {
      var analyticsSheet = getSheetByName(allSheets, SHEETS_TITLES.ANALYTICKS);

      if (analyticsSheet && analyticsSheet.getName() != activeSheet.getName())
      {
        var currentRow = cell.getRow();
        var prjName = activeSheet.getRange(currentRow, getColumnNumberByName(activeSheet, INITIAL_COLUMNS_TITLES.TASK_NAME)).getValue().toString();
        var prjDesc = activeSheet.getRange(currentRow, getColumnNumberByName(activeSheet, INITIAL_COLUMNS_TITLES.TASK_DESCRIPTION)).getValue().toString();

        if (!findSameProjectInOtherSheet(analyticsSheet, prjName))
        {
          var lastRow = getLastFreeRow(analyticsSheet);
          copyCellData(activeSheet, analyticsSheet, currentRow, INITIAL_COLUMNS_TITLES.TASK_NAME, lastRow, OUTPUT_COLUMNS_TITLES.TASK_NAME);
          copyCellData(activeSheet, analyticsSheet, currentRow, 'null', lastRow, OUTPUT_COLUMNS_TITLES.TASK_INITIATOR, convertInitiatorFromSheetName(activeSheet));
          copyCellData(activeSheet, analyticsSheet, currentRow, INITIAL_COLUMNS_TITLES.TASK_DESCRIPTION, lastRow, OUTPUT_COLUMNS_TITLES.TASK_DESCRIPTION);
          copyCellData(activeSheet, analyticsSheet, currentRow, 'null', lastRow, OUTPUT_COLUMNS_TITLES.TASK_STATUS, 'Backlog');


          console.log("Mailing: ", convertInitiatorFromSheetName(activeSheet), prjName, prjDesc, lastRow.toString());

          MailApp.sendEmail
          ({
            to: getMailAdresses(getSheetByName(allSheets, SHEETS_TITLES.PICKERS), 'analitycs'),
            subject: "New task in backlog", 
            htmlBody: 
            (
              "Новая задача от: <p style='color:rgba(255,0,0,0.5);'>" + convertInitiatorFromSheetName(activeSheet) + "</p><br>" +
              "Название задачи: <p style='color:rgba(255,0,0,0.5);'>" + prjName + "</p><br>" +
              "Описание задачи: <p style='color:rgba(255,0,0,0.5);'>" + prjDesc + "</p><br>" +
              "Номер строки в беклоге: <p style='color:rgba(255,0,0,0.5);'>" + lastRow.toString() + "</p><br>"
            )
          });
        }
      }
    }

    if (cellValue == 'Design / Product')
    {
      // TODO
    }
  }
}
