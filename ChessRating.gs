var highlight = '#ff9900';
var name_firstlast = '';
var name_lastfirst = '';
var regular_rating = 0;
var quick_rating = 0;
var blitz_rating = 0;
var highest_rating = 0;
var note = null;

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
  
  // reset background
  var range = cur.getDataRange();
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  for (var i = 1; i <= rows; i++) {
    for (var j = 1; j <= cols; j++) {
      var cell = range.getCell(i,j);
      var color = cell.getBackground();
      if (color == highlight)
        cell.setBackground('white');
    }
  }
  SpreadsheetApp.flush();
}

function UpdateCell(cell, value, mode) {
  if (value) {
    var old = cell.getValue();
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
    cell.setValue(value);
    cell.setBackground(highlight);
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
  var html = UrlFetchApp.fetch('http://www.uschess.org/msa/MbrDtlMain.php?' + id).getContentText();
  
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
  var html = UrlFetchApp.fetch('http://chess.ca/players?check_rating_number=' + id).getContentText();
  
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
  var html = UrlFetchApp.fetch('http://ratings.fide.com/card.phtml?event=' + id).getContentText();
   
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

function UpdateOneRow(range, i) {
  var org = range.getCell(i,1).getValue();
  var id = range.getCell(i,2).getValue();
  name_firstlast = '';
  name_lastfirst = '';
  regular_rating = 0;
  quick_rating = 0;
  blitz_rating = 0;
  highest_rating = 0;
  note = null;
  
  if (org == 'USCF' && id) {
    SearchUSCF(id);
  } else if (org == 'CFC' && id) {
    SearchCFC(id);
  } else if (org == 'FIDE' && id) {
    SearchFIDE(id);
  } else if (org != '') {
    name_firstlast = range.getCell(i,3).getValue();
    name_lastfirst = range.getCell(i,4).getValue();
    note = 'Unsupported Org';
  }
  
  if (name_lastfirst == '' && name_firstlast != '')
    name_lastfirst = LastFirstName(name_firstlast);
  else if (name_firstlast == '' && name_lastfirst != '')
    name_firstlast = FirstLastName(name_lastfirst);
  
  var update = false;
  update |= UpdateCell(range.getCell(i, 3), name_firstlast, 0);
  update |= UpdateCell(range.getCell(i, 4), name_lastfirst, 0);
  update |= UpdateCell(range.getCell(i, 6), regular_rating, 1);
  update |= UpdateCell(range.getCell(i, 7), quick_rating, 1);
  update |= UpdateCell(range.getCell(i, 8), blitz_rating, 1);
  update |= UpdateCell(range.getCell(i, 5), highest_rating, 2);
  update |= UpdateCell(range.getCell(i, 9), note, 0);
  return update;
}

function UpdateSelectedRows() {
  // handle multiple users
  var lock = LockService.getScriptLock();
  var success = lock.tryLock(1);
  if (!success) {
    Browser.msgBox('Another user is updating the sheet, please try again later.');
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cur = ss.getSheetByName('Current');
  var range = cur.getDataRange();
  var rows = cur.getActiveRange();
  var n = 0;
  for (var i = rows.getRow(); i <= rows.getLastRow(); i++) {
    if (UpdateOneRow(range, i))
      n ++;
  }
  
  lock.releaseLock();
  Browser.msgBox(n + ' persons updated');
}

function UpdateAllRows() {
  // handle multiple users
  var lock = LockService.getScriptLock();
  var success = lock.tryLock(1);
  if (!success) {
    Browser.msgBox('Another user is updating the sheet, please try again later.');
    return;
  }
  
  // save a backup first
  BackupSheet();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cur = ss.getSheetByName('Current');
  var range = cur.getDataRange();
  var rows = range.getNumRows();
  var n = 0;
  for (var i = 6; i <= rows; i++) {
    if (UpdateOneRow(range, i))
      n ++;
  }
  
  lock.releaseLock();
  Browser.msgBox(n + ' persons updated');
}

