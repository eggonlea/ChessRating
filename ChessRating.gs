var subject = '[ChessRating] update notification';
var body = '';

var highlight = '#ff9900';
var startRow = 5;
var cols = 10;

var err_network = 'Network ERR';
var err_fed = 'Unsupported FED';
var err_id = 'Invalid ID';

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
  ss.moveActiveSheet(4);
  ss.setActiveSheet(cur);
  
  SpreadsheetApp.flush();
}

function IsError(value) {
  if (value == err_network)
    return true;
  if (value == err_fed)
    return true;
  if (value == err_id)
    return true;
  return false;
}

function UpdateCell(values, colors, i, j, value, mode) {
  var old = values[i][j];
  
  // deal with erro note specifically
  if (mode == 3) {
    if (IsError(old) && old != value)
      values[i][j] = ''; // clear previous error note
    else if (old != '')
      return false; // keep manually inputted note
  }
  
  // continue normal process
  if (value != null) {
    switch (mode) {
      case 3: // note
        break;
      case 2: // max rating
        if (value <= old)
          return false;
        break;
      case 1: // rating
        highest_rating = Math.max(value, highest_rating);
        if (value == old)
          return false;
        break;
      case 0: // name/link
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

function FixNameAddComma(name) {
  if (!name)
    return null;
  
  var last = name.split(' ').slice(0, 1).join(' ').trim();
  var first = name.split(' ').slice(1).join(' ').trim();
  return last + ', ' + first;
}

function SearchUSCF(id) {
  var html = null;
  link = 'http://www.uschess.org/msa/MbrDtlRtgSupp.php?' + id;
  try {
    html = UrlFetchApp.fetch(link).getContentText();
  } catch (err) {
    html = null;
  }
  
  if (!html) {
    note = err_network;
    return;
  }
  
  // <font size=+1><b>15230922: JUNREN LI</b></font>
  var ret = html.match('<b>' + id + ': (.+)</b>');
  if (ret)
    name_firstlast = ret[1];
  else
    note = err_id;
  
  // <tr bgcolor=FFFFC0 align=center>
  // <td width=200>&nbsp;</td>
  // <td width=100>2014-03</td>
  // <td width=100> 107 (P05)</td>
  // <td width=100> 114 (P05)</td>
  // <td width=100>---</td>
  // <td width=100>---</td>
  // <td width=100>---</td>
  // <td></td>
  var re = /<td width=100>\d+-\d+<\/td>\n<td width=100>(?:(?: *(\d+)(?: \(P\d+\))*)|(?:---))<\/td>\n<td width=100>(?:(?: *(\d+)(?: \(P\d+\))*)|(?:---))<\/td>\n<td width=100>(?:(?: *(\d+)(?: \(P\d+\))*)|(?:---))<\/td>\n/g;
  while ((ret = re.exec(html)) != null) {
    var rating = Number(ret[1]);
    if (rating) {
      if (regular_rating == 0)
        regular_rating = rating;
      highest_rating = Math.max(rating, highest_rating);
    }
    rating = Number(ret[2]);
    if (rating) {
      if (quick_rating == 0)
        quick_rating = rating;
      highest_rating = Math.max(rating, highest_rating);
    }
    rating = Number(ret[3]);
    if (rating) {
      if (blitz_rating == 0)
        blitz_rating = rating;
      highest_rating = Math.max(rating, highest_rating);
    }
  }
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
    note = err_network;
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
    note = err_id;
  
  // City/Prov</td></tr><tr><td>1286</td><td>1286</td><td>1225</td><td>1225</td>
  var ret = html.match('City/Prov</td></tr><tr><td>([0-9]+)</td><td>([0-9]+)</td><td>([0-9]+)</td><td>([0-9]+)</td>');
  if (ret) {
    regular_rating = Number(ret[1]);
    quick_rating = Number(ret[3]);
    highest_rating = Math.max(Number(ret[2]), highest_rating);
    highest_rating = Math.max(Number(ret[4]), highest_rating);
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
    note = err_network;
    return;
  }
  
  // <title>Hu, Zheng  FIDE Chess Profile
  var ret = html.match('<title>(.+) *FIDE Chess Profile');
  if (ret)
    name_lastfirst = ret[1];
  else
    note = err_id;
  
  // <small>std.</small><br>1318
  var ret = html.match('<small>std.</small><br>([0-9]+)');
  if (ret)
    regular_rating = Number(ret[1]);

  // <small>rapid</small><br>Not rated
  var ret = html.match('<small>rapid</small><br>([0-9]+)');
  if (ret)
    quick_rating = Number(ret[1]);

  // <small>blitz</small><br>Not rated
  var ret = html.match('<small>blitz</small><br>([0-9]+)');
  if (ret)
    blitz_rating = Number(ret[1]);
}

function SearchCMA(id) {
  var html = null;
  link = 'https://chess-math.org/cotes/id/' + id;
  try {
    html = UrlFetchApp.fetch(link).getContentText();
  } catch (err) {
    html = null;
  }
  
  if (!html) {
    note = err_network;
    return;
  }
  
  // <h4><spam style="color:orange">Li Lang Ji<span>
  var ret = html.match('<h4><spam.*>(.+)<span>');
  if (ret)
    name_lastfirst = FixNameAddComma(ret[1]);
  else
    note = err_id;
  
  // <h4>Rating : 484</h4>
  var ret = html.match('<h4>Rating : ([0-9]+)</h4>');
  if (ret)
    regular_rating = Number(ret[1]);

  // <h4>Max rating : 484</h4>
  var ret = html.match('<h4>Max rating : ([0-9]+)</h4>');
  if (ret)
    highest_rating = Math.max(Number(ret[1]), highest_rating);
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
  } else if (fed == 'CMA' && id) {
    SearchCMA(id);
  } else if (fed != '') {
    name_firstlast = values[i][2];
    name_lastfirst = values[i][3];
    note = err_fed;
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
  update |= UpdateCell(values, colors, i, 8, note, 3);
  update |= UpdateCell(values, colors, i, 9, link, 0);
  
  if (update)
    body += '[' + (i + startRow) + ']' + values[i] + '\n';
  
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
  var row1 = Math.max(startRow, selected.getRow());
  var row2 = Math.min(cur.getLastRow(), selected.getLastRow());
  if (row2 < row1) {
    Browser.msgBox('Please choose any rows starting from ROW ' + startRow);
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
  if (row1 > startRow || row2 > startRow)
    MailApp.sendEmail(ss.getEditors(), subject, body);
  
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
  var row1 = startRow;
  var row2 = cur.getLastRow();
  if (row2 < row1) {
    Browser.msgBox('Please input data starting from ROW ' + startRow);
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
  MailApp.sendEmail(ss.getEditors(), subject, body);
  
  return n;
}

function UpdateLog(msg) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('UpdateLog');
  var time = Utilities.formatDate(new Date(), 'GMT-7', 'EEE yyyy-MM-dd HH:mm:ss z');
  sheet.appendRow([time, msg]);
  body += '\n' + time + ": " + msg + '\n\n';
}

// PST 1-2am everyday
function AutoUpdate() {
  UpdateLog('Triggering auto update...');
  UpdateAllRows();
}
