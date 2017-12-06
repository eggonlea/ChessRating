var highlight = '#ff9900';
var startRow1 = 5;
var startRow2 = 6;
var cols = 10;

var name_firstlast = '';
var name_lastfirst = '';
var regular_rating = 0;
var quick_rating = 0;
var blitz_rating = 0;
var highest_rating = 0;
var note = null;
var link = null;

function BackupSheet() {
  var bak = 'bak.' + Utilities.formatDate(new Date(), 'GMT-7', 'yyyy-MM-dd');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // delete old backup sheet
  var old = ss.getSheetByName(bak);
  if (old)
    ss.deleteSheet(old);
  
  // make a new copy of current sheet
  var cur = ss.getSheetByName('Current');
  var sheet = cur.copyTo(ss);
  sheet.setName(bak);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(2);
  ss.setActiveSheet(cur);
  
  SpreadsheetApp.flush();
}

function UpdateCell(values, colors, i, j, value, mode) {
  if (value) {
    var old = values[i][j];
    switch (mode) {
      case 2: // max rating
        if (value <= old)
          return false;
        break;
      case 1: // rating
        if (value > highest_rating)
          highest_rating = value;
        if (value == old)
          return false;
        break;
      case 0: // name
      default: // should never reach here
        if (value == old)
          return false;
        break;
    }
    values[i][j] = value;
    colors[i][j] = highlight;
    return true;
  }
  return false;
}

function FirstLastName(name) {
  if (!name)
    return null;
  
  var last = name.split(',').slice(0, -1).join('').trim();
  var first = name.split(',').slice(-1).join('').trim();
  return first + ' ' + last;
}

function LastFirstName(name) {
  if (!name)
    return null;
  
  var first = name.split(' ').slice(0, -1).join(' ').trim();
  var last = name.split(' ').slice(-1).join(' ').trim();
  return last + ', ' + first;
}

function SearchUSCF(id) {
  var html = null;
  link = 'http://www.uschess.org/msa/MbrDtlMain.php?' + id;
  try {
    html = UrlFetchApp.fetch(link).getContentText();
  } catch (err) {
    html = null;
  }
  
  if (!html) {
    note = 'Network Err';
    return;
  }
  
  // <font size=+1><b>15230922: JUNREN LI</b></font>
  var ret = html.match('<b>' + id + ': (.+)</b>');
  if (ret)
    name_firstlast = ret[1];
  else
    note = 'Invalid ID';
  
  // Regular Rating
  // </td>
  //
  // <td>
  // <b><nobr>
  // 660&nbsp;&nbsp;
  // 2017-07</nobr>
  var ret = html.match('Regular Rating\n</td>\n+<td>\n<b><nobr>\n([0-9]+)');
  if (ret)
    regular_rating = ret[1];
  
  // Quick Rating
  // </td>
  //
  // <td>
  // <b>
  // 658&nbsp;&nbsp;
  // 2017-07</nobr>
  var ret = html.match('Quick Rating\n</td>\n+<td>\n<b>\n([0-9]+)');
  if (ret)
    quick_rating = ret[1];
  
  // Blitz Rating
  // </td>
  //
  // <td>
  // <b>
  // (Unrated)&nbsp;&nbsp;
  // 2017-07</nobr>
  var ret = html.match('Blitz Rating\n</td>\n+<td>\n<b>\n([0-9]+)');
  if (ret)
    blitz_rating = ret[1];
}

function SearchCFC(id) {
  var html = null;
  link = 'http://chess.ca/players?check_rating_number=' + id;
  try {
    html = UrlFetchApp.fetch(link).getContentText();
  } catch (err) {
    html = null;
  }
  
  if (!html) {
    note = 'Network Err';
    return;
  }
  
  // <h2>Player Information</h2>
  //  
  //   <div class='content'>
  //     <table class='tbl_player_results'><thead><tr><th colspan='8'>Adam Li</th>
  var ret = html.match('<h2>Player Information</h2>\n *\n *<div.*>\n *<table.*><thead><tr><th.*>(.+)</th>');
  if (ret)
    name_firstlast = ret[1];
  else
    note = 'Invalid ID';
  
  // City/Prov</td></tr><tr><td>1286</td><td>1286</td><td>1225</td><td>1225</td>
  var ret = html.match('City/Prov</td></tr><tr><td>([0-9]+)</td><td>([0-9]+)</td><td>([0-9]+)</td><td>([0-9]+)</td>');
  if (ret) {
    regular_rating = ret[1];
    quick_rating = ret[3];
    if (ret[2] > highest_rating)
      highest_rating = ret[2];
    if (ret[4] > highest_rating)
      highest_rating = ret[4];
  }
}

function SearchFIDE(id) {
  var html = null;
  link = 'http://ratings.fide.com/card.phtml?event=' + id;
  try {
    html = UrlFetchApp.fetch(link).getContentText();
  } catch (err) {
    html = null;
  }
  
  if (!html) {
    note = 'Network Err';
    return;
  }
  
  // <title>Hu, Zheng  FIDE Chess Profile
  var ret = html.match('<title>(.+) *FIDE Chess Profile');
  if (ret)
    name_lastfirst = ret[1];
  else
    note = 'Invalid ID';
  
  // <small>std.</small><br>1318
  var ret = html.match('<small>std.</small><br>([0-9]+)');
  if (ret)
    regular_rating = ret[1];

  // <small>rapid</small><br>Not rated
  var ret = html.match('<small>rapid</small><br>([0-9]+)');
  if (ret)
    quick_rating = ret[1];

  // <small>blitz</small><br>Not rated
  var ret = html.match('<small>blitz</small><br>([0-9]+)');
  if (ret)
    blitz_rating = ret[1];
}

function UpdateOneRow(values, colors, i) {
  var fed = values[i][0];
  var id = values[i][1];
  name_firstlast = '';
  name_lastfirst = '';
  regular_rating = 0;
  quick_rating = 0;
  blitz_rating = 0;
  highest_rating = 0;
  note = null;
  link = null;
  
  if (fed == 'USCF' && id) {
    SearchUSCF(id);
  } else if (fed == 'CFC' && id) {
    SearchCFC(id);
  } else if (fed == 'FIDE' && id) {
    SearchFIDE(id);
  } else if (fed != '') {
    name_firstlast = values[i][2];
    name_lastfirst = values[i][3];
    note = 'Unsupported fed';
  }
  
  if (name_lastfirst == '' && name_firstlast != '')
    name_lastfirst = LastFirstName(name_firstlast);
  else if (name_firstlast == '' && name_lastfirst != '')
    name_firstlast = FirstLastName(name_lastfirst);
  
  var update = false;
  update |= UpdateCell(values, colors, i, 2, name_firstlast, 0);
  update |= UpdateCell(values, colors, i, 3, name_lastfirst, 0);
  update |= UpdateCell(values, colors, i, 5, regular_rating, 1);
  update |= UpdateCell(values, colors, i, 6, quick_rating, 1);
  update |= UpdateCell(values, colors, i, 7, blitz_rating, 1);
  update |= UpdateCell(values, colors, i, 4, highest_rating, 2);
  update |= UpdateCell(values, colors, i, 8, note, 0);
  update |= UpdateCell(values, colors, i, 9, link, 0);
  return update;
}

function UpdateSelectedRows() {
  // handle multiple users
  var lock = LockService.getScriptLock();
  var success = lock.tryLock(1);
  if (!success) {
    Browser.msgBox('Another user is updating the sheet, please try again later.');
    return 0;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cur = ss.getSheetByName('Current');
  var selected = cur.getActiveRange();
  var row1 = selected.getRow();
  var row2 = selected.getLastRow();
  if (row1 < startRow1)
    row1 = startRow1;
  if (row2 < row1) {
    Browser.msgBox('Please choose any rows starting from ROW 8');
    return 0;
  }
  
  var rows = row2 - row1 + 1;
  var range = cur.getRange(row1, 1, rows, cols);
  var values = range.getValues();
  var colors = range.getBackgrounds();

  for (var i = 0; i < rows; i++)
    for (var j = 2; j < cols; j++)
      colors[i][j] = '#ffffff';
  
  range.setBackgrounds(colors);
  
  SpreadsheetApp.flush();
  
  var n = 0;
  for (var i = 0; i < rows; i++) {
    if (UpdateOneRow(values, colors, i))
      n ++;
  }
  
  range.setValues(values);
  range.setBackgrounds(colors);
  SpreadsheetApp.flush();
  lock.releaseLock();
  Browser.msgBox(n + ' persons updated');
  UpdateLog('UpdateSelectedRows [' + row1 + '-' + row2 + ']: ' + n + ' person(s) updated');
  
  return n;
}

function UpdateAllRows() {
  // handle multiple users
  var lock = LockService.getScriptLock();
  var success = lock.tryLock(1);
  if (!success) {
    Browser.msgBox('Another user is updating the sheet, please try again later.');
    return 0;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cur = ss.getSheetByName('Current');
  var row1 = startRow2;
  var row2 = cur.getLastRow();
  if (row2 < row1) {
    Browser.msgBox('Please input data starting from ROW 8');
    return 0;
  }
  
  // save a backup first
  BackupSheet();
  
  var rows = row2 - row1 + 1;
  var range = cur.getRange(row1, 1, rows, cols);
  var values = range.getValues();
  var colors = range.getBackgrounds();

  for (var i = 0; i < rows; i++)
    for (var j = 2; j < cols; j++)
      colors[i][j] = '#ffffff';
  
  range.setBackgrounds(colors);
  
  SpreadsheetApp.flush();
  
  var n = 0;
  for (var i = 0; i < rows; i++) {
    if (UpdateOneRow(values, colors, i))
      n ++;
  }
  
  range.setValues(values);
  range.setBackgrounds(colors);
  SpreadsheetApp.flush();
  lock.releaseLock();
  try {
    Browser.msgBox(n + ' persons updated');
  } catch (err) {
  }
  
  UpdateLog('UpdateAllRows [' + row1 + '-' + row2 + ']: ' + n + ' person(s) updated');
              
  return n;
}

function UpdateLog(msg) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UpdateLog');
  var time = Utilities.formatDate(new Date(), 'GMT-7', 'yyyy-MM-dd hh:mm:ss');
  sheet.appendRow([time, msg]);
}

// PST 1-2am everyday
function AutoUpdate() {
  UpdateLog('Triggering auto update...');
  UpdateAllRows();
}
