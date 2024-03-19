// function createTimeDrivenTriggers()
// {
//   ScriptApp.newTrigger('checkDeadLinesDverdue')
//       .timeBased()
//       .atHour(9)
//       .create();
// }

function getSheetUrl(ss, sheet)
{
  var url = '';
  url += ss.getUrl();
  url += '#gid=';
  url += sheet.getSheetId(); 

  return url;
}


const SHEETS_TITLES =
{
  DEMAND:     'Demand Backlog',
  SUPPLY:     'Supply Backlog',
  BALANCE:    'Balance Backlog',
  EXECUTION:  'NONE',
  ANALYTICKS: 'Analytics backlog',
  PRODUCT:    'Product backlog',
  PPL_OPS:    'Common ppl ops',
  // PICKERS:    'Pickers -->',
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
            to: getMailAdresses(getSheetByName(allSheets, 'Pickers -->'/*SHEETS_TITLES.PICKERS*/), 'analitycs'),
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

function checkDeadLinesDverdue()
{
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = activeSpreadsheet.getSheets();

  var pickersSheet = getSheetByName(allSheets, 'Pickers -->');

  if (!pickersSheet)
  {
    console.log("ERROR: checkDeadLinesDverdue: !pickersSheet", pickersSheet);
    return null;
  }

  var pplNamesCol = getColumnNumberByName(pickersSheet, 'Имя', 3);
  var pplEmailsCol = getColumnNumberByName(pickersSheet, 'Почта', 3);

  if (!pplNamesCol || !pplEmailsCol)
  {
    console.log("ERROR: checkDeadLinesDverdue: !pplNamesCol || !pplEmailsCol", pplNamesCol, pplEmailsCol);
    return null;
  }
	
  var pplStr = []; // {name, email}
  
  for (var i = 4; i < pickersSheet.getMaxRows(); ++i)
  {
    var name = pickersSheet.getRange(i, pplNamesCol).getValue();
    var email = pickersSheet.getRange(i, pplEmailsCol).getValue();

    if (!(name.length === 0) && !(email.length === 0))
      pplStr.push({name: name, email: email});
  }

  if (pplStr.length == 0)
  {
    console.log("ERROR: checkDeadLinesDverdue: pplStr.length == 0");
    return null;
  }

  console.log(pplStr);

  var neededSheets = Object.values(SHEETS_TITLES);

  for (var sheetIdx = 0; sheetIdx < neededSheets.length; ++sheetIdx)
  {
    var sheet = getSheetByName(allSheets, neededSheets[sheetIdx]);

    if (sheet)
    {
      var counter = 0;
      console.log(sheet.getName());

      var taskCol   = getColumnNumberByName(sheet, INITIAL_COLUMNS_TITLES.TASK_NAME);
      var ownerCol  = getColumnNumberByName(sheet, INITIAL_COLUMNS_TITLES.TASK_OWNER);
      var statusCol = getColumnNumberByName(sheet, INITIAL_COLUMNS_TITLES.TASK_STATUS);
      var etaCol    = getColumnNumberByName(sheet, INITIAL_COLUMNS_TITLES.TASK_DEADLINE);
      
      if (etaCol && ownerCol && statusCol && taskCol)
      {
        var lastRow = getLastFreeRow(sheet);
        var currentDate = new Date();
        var sendList = [];

        for (var i = 3; i < lastRow; ++i)
        {
          var eta = new Date(sheet.getRange(i, etaCol).getValue());
          var status = sheet.getRange(i, statusCol).getValue();
        
          if (eta && Object.prototype.toString.call(eta) === '[object Date]')
          {
            var diffDays = Math.ceil((currentDate - eta) / (1000 * 60 * 60 * 24)); 

            if (diffDays > 1 && !(['Done - success', 'On hold', 'Not doing', 'Done - failure'].includes(status)))
            {
              counter++;
              var owner = sheet.getRange(i, ownerCol).getValue();

              if (!(owner.length === 0))
              {
                var manIndex = pplStr.findIndex((x) => {return x.name == owner});
                if (manIndex != -1)
                {
                  var searchingIdx = sendList.findIndex((x) => {return x.email == pplStr[manIndex].email});
                  var body =
                  (
                    sheet.getRange(i, taskCol).getValue().toString() + "<br>"
                    + '<p style="color:rgba(255,0,0,0.5);"">'
                    + "Номер строки в беклоге: " + i.toString() + "<br>"
                    + "Статус задачи: " + status.toString() + "<br>"
                    + "Просрочка: " + diffDays.toString() + " дней"
                    + "</p><br>"
                  );

                  if (searchingIdx > -1)
                    sendList[searchingIdx].body += body;
                  else
                    sendList.push
                    ({
                      email: pplStr[manIndex].email, 
                      body: ('<a href="' + getSheetUrl(activeSpreadsheet, sheet) + '">Ссылка на Master backlog</a><br><br>' + body)
                    });
                }
              }
            }
          }
        }

        console.log(sendList);
        console.log(sheet.getName(), counter);

        for (var j = 0; j < sendList.length; ++j)
        {
          MailApp.sendEmail
          ({
            to: sendList[j].email,
            subject: "Deadline OVERDUE  in Master backlog", 
            htmlBody: sendList[j].body
          });
        }
      }
    }
  }
}







