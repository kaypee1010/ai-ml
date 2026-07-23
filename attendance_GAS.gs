// ================================================================
// ATTENDANCE SYSTEM — Google Apps Script (v2 with RI checks)
// Paste entire file in Apps Script editor, redeploy as new version
// ================================================================

const SESSION_STORE_KEY = 'active_sessions';

// Internship list — populate when KP shares
const INTERNSHIP_STUDENTS = [];


function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'session_start')  return handleSessionStart(data);
    if (data.type === 'session_close')  return handleSessionClose(data);
    if (data.type === 'attendance')     return handleAttendance(data);
    return jsonResponse({success:false, error:'Unknown request type'});
  } catch(err) {
    return jsonResponse({success:false, error: err.toString()});
  }
}
function doGet(e) { return doPost(e); }

// ── Session Start ──────────────────────────────────────────────
function handleSessionStart(data) {
  const key = data.subject + '_' + data.section;
  const expiry = Date.now() + (7 * 60 * 1000);
  const store = getStore();
  store[key] = { otp: data.otp, expiry, date: data.date };
  setStore(store);
  ensureSheet(data.subject, data.section, data.date);
  return jsonResponse({success:true});
}

// ── Session Close ──────────────────────────────────────────────
function handleSessionClose(data) {
  const key = data.subject + '_' + data.section;
  const store = getStore();
  delete store[key];
  setStore(store);
  return jsonResponse({success:true});
}

// ── Attendance with full Referential Integrity ─────────────────
function handleAttendance(data) {

  // ── RI CHECK 1: reg no must exist in master roster ──────────
  // ── RI CHECK 0: internship block ────────────────────────────
  if (INTERNSHIP_STUDENTS.includes(data.reg.toString().trim())) {
    return jsonResponse({success:false, error:'You are marked on internship exemption. Attendance is not applicable for your record this semester. Contact Dr. Kamalpreet Singh if this is incorrect.'});
  }

  const masterMatch = findStudentInMaster(data.subject, data.reg);
  if (!masterMatch) {
    return jsonResponse({
      success: false,
      error: 'Registration number ' + data.reg + ' not found in the master roster for ' + data.subject + '. Submission rejected.'
    });
  }

  // ── RI CHECK 2: section from browser must match roster ───────
  // masterMatch.section is the authoritative section from our roster
  if (masterMatch.section !== data.section) {
    return jsonResponse({
      success: false,
      error: 'Section mismatch. Your registration ' + data.reg + ' belongs to Section ' +
             masterMatch.section + ', not Section ' + data.section +
             '. Please select Section ' + masterMatch.section + ' and try again.'
    });
  }

  // ── RI CHECK 3: name from browser must match roster ──────────
  // (case-insensitive — student can't submit someone else's name)
  if (masterMatch.name.toLowerCase() !== data.name.toLowerCase()) {
    return jsonResponse({
      success: false,
      error: 'Name mismatch for reg ' + data.reg + '. Expected: ' + masterMatch.name + '. Submission rejected.'
    });
  }

  // ── CHECK 4: active session exists for this subject+section ──
  const key = data.subject + '_' + masterMatch.section;
  const store = getStore();
  const session = store[key];
  if (!session) {
    return jsonResponse({success:false, error:'No active attendance session open for ' + data.subject + ' Section ' + masterMatch.section + '.'});
  }

  // ── CHECK 5: OTP not expired ─────────────────────────────────
  if (Date.now() > session.expiry) {
    delete store[key]; setStore(store);
    return jsonResponse({success:false, error:'OTP has expired. Attendance window is now closed.'});
  }

  // ── CHECK 6: OTP correct ─────────────────────────────────────
  if (data.otp !== session.otp) {
    return jsonResponse({success:false, error:'Incorrect OTP. Please check the code shown on the projector.'});
  }

  // ── CHECK 7: duplicate submission ────────────────────────────
  const sheetName = data.subject + '_' + masterMatch.section;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ensureSheet(data.subject, masterMatch.section, session.date);

  const dateCol = getOrCreateDateColumn(sheet, session.date);
  const studentRow = findStudentRow(sheet, data.reg);

  if (studentRow === -1) {
    return jsonResponse({success:false, error:'Student row not found in sheet. Contact faculty.'});
  }

  const existing = sheet.getRange(studentRow, dateCol).getValue();
  if (existing === 'P' || existing === 'P⚠') {
    return jsonResponse({success:false, error:'Attendance already marked for today. Duplicate submission blocked.'});
  }

  // ── ALL CHECKS PASSED — mark attendance ──────────────────────
  const mark = (data.gpsStatus === 'on-campus') ? 'P' : 'P⚠';
  sheet.getRange(studentRow, dateCol).setValue(mark);

  // Color code: green for P, amber for P⚠
  sheet.getRange(studentRow, dateCol)
       .setBackground(mark === 'P' ? '#c6efce' : '#ffeb9c')
       .setFontColor(mark === 'P' ? '#276221' : '#9c5700');

  logGPSEntry(data, mark, masterMatch.section);

  return jsonResponse({
    success: true,
    message: 'Attendance marked (' + mark + ') for ' + masterMatch.name + ', Section ' + masterMatch.section
  });
}

// ── Find student across all sections of a subject ──────────────
// Returns {reg, name, section} or null
function findStudentInMaster(subject, reg) {
  const allRosters = getRoster(subject, 'ALL');
  for (const section in allRosters) {
    const found = allRosters[section].find(s => s.reg === reg.toString().trim());
    if (found) return { reg: found.reg, name: found.name, section: section };
  }
  return null;
}

// ── Find student row in sheet by reg no ───────────────────────
function findStudentRow(sheet, reg) {
  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const val = sheet.getRange(r, 2).getValue().toString().trim();
    if (val === reg.toString().trim()) return r;
  }
  return -1;
}

// ── Ensure sheet exists with roster pre-populated ─────────────
function ensureSheet(subject, section, date) {
  const sheetName = subject + '_' + section;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Header row
    sheet.getRange(1,1,1,4).setValues([['S.No','Reg No','Name','Section']]);
    sheet.getRange(1,1,1,4)
         .setBackground('#8B3A3A').setFontColor('#fff').setFontWeight('bold');

    // Populate roster
    const roster = getRoster(subject, section);
    roster.forEach((s, i) => {
      sheet.getRange(i+2, 1).setValue(i+1);
      sheet.getRange(i+2, 2).setValue(s.reg);
      sheet.getRange(i+2, 3).setValue(s.name);
      sheet.getRange(i+2, 4).setValue(section);
    });

    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(4);
    sheet.autoResizeColumns(1, 4);
  }

  getOrCreateDateColumn(sheet, date);
  return sheet;
}

// ── Get or create a date column, pre-fill 'A' for all students ─
function getOrCreateDateColumn(sheet, date) {
  const lastCol = Math.max(sheet.getLastColumn(), 4);
  for (let c = 5; c <= lastCol; c++) {
    if (sheet.getRange(1, c).getValue().toString() === date) return c;
  }
  // New date column
  const newCol = lastCol + 1;
  const header = sheet.getRange(1, newCol);
  header.setValue(date).setBackground('#8B3A3A').setFontColor('#fff').setFontWeight('bold');

  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    sheet.getRange(r, newCol).setValue('A')
         .setBackground('#ffc7ce').setFontColor('#9c0006'); // red for absent
  }
  sheet.autoResizeColumn(newCol);
  return newCol;
}

// ── GPS log tab ───────────────────────────────────────────────
function logGPSEntry(data, mark, verifiedSection) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName('GPS_Log');
  if (!log) {
    log = ss.insertSheet('GPS_Log');
    log.appendRow(['Timestamp','Subject','Section (browser)','Section (verified)',
                   'Reg','Name','Date','Time','OTP_Valid','GPS_Status','Lat','Lng','Accuracy_m','Mark']);
    log.getRange(1,1,1,14).setBackground('#8B3A3A').setFontColor('#fff').setFontWeight('bold');
  }
  log.appendRow([
    new Date().toLocaleString('en-IN'),
    data.subject,
    data.section,        // what browser sent
    verifiedSection,     // what roster confirmed
    data.reg, data.name,
    data.date, data.time,
    'YES',
    data.gpsStatus,
    data.gpsLat, data.gpsLng, data.gpsAccuracy,
    mark
  ]);
}

// ── Script Properties helpers ─────────────────────────────────
function getStore() {
  const raw = PropertiesService.getScriptProperties().getProperty(SESSION_STORE_KEY);
  return raw ? JSON.parse(raw) : {};
}
function setStore(obj) {
  PropertiesService.getScriptProperties().setProperty(SESSION_STORE_KEY, JSON.stringify(obj));
}
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
         .setMimeType(ContentService.MimeType.JSON);
}

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Attendance Tools')
    .addItem('Convert All → SGT Format', 'convertToSGTFormat')
    .addSeparator()
    .addItem('Convert AI_ML_A → SGT', function(){ convertSingleSection('AI_ML_A'); })
    .addItem('Convert AI_ML_B → SGT', function(){ convertSingleSection('AI_ML_B'); })
    .addItem('Convert AI_ML_C → SGT', function(){ convertSingleSection('AI_ML_C'); })
    .addItem('Convert AI_ML_D → SGT', function(){ convertSingleSection('AI_ML_D'); })
    .addItem('Convert MATLAB_C → SGT',  function(){ convertSingleSection('MATLAB_C'); })
    .addToUi();
}

// ── SGT Converter ─────────────────────────────────────────────
function convertToSGTFormat() {
  const sheets = ['AI_ML_A','AI_ML_B','AI_ML_C','AI_ML_D','MATLAB_C'];
  sheets.forEach(s => convertSingleSection(s));
  SpreadsheetApp.getUi().alert('✅ Done! SGT format sheets created.');
}

function convertSingleSection(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(sheetName);
  if (!src) return;

  const outName = sheetName + '_SGT';
  const old = ss.getSheetByName(outName);
  if (old) ss.deleteSheet(old);
  const out = ss.insertSheet(outName);

  const data = src.getDataRange().getValues();
  const headers = data[0];
  const DATE_START = 4;

  out.getRange(1,1,1,headers.length).setValues([headers]);
  out.getRange(1,1,1,headers.length).setBackground('#8B3A3A').setFontColor('#fff').setFontWeight('bold');

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const outRow = row.slice(0, DATE_START);
    // Internship students stay 0 throughout — EX doesn't count
    const regNo = row[1] ? row[1].toString().trim() : '';
    const isInternship = INTERNSHIP_STUDENTS.includes(regNo);
    let count = 0;
    for (let c = DATE_START; c < headers.length; c++) {
      const v = row[c];
      if (!isInternship && (v === 'P' || v === 'P⚠')) count++;
      outRow.push(isInternship ? 0 : count);
    }
    out.getRange(r+1, 1, 1, outRow.length).setValues([outRow]);
  }

  out.setFrozenRows(1);
  out.setFrozenColumns(4);
  out.autoResizeColumns(1, headers.length);
}

// ── Master Roster ─────────────────────────────────────────────
// getRoster(subject, 'ALL') returns {A:[...], B:[...], C:[...], D:[...]}
// getRoster(subject, 'A')   returns just section A array
function getRoster(subject, section) {
  const R = {
    'AI_ML': {
      'A': [{reg:'231302002',name:'Tannu Tiwari'},{reg:'231302003',name:'Pranav Singh'},{reg:'231302006',name:'Tanish'},{reg:'231302007',name:'Prince Kumar'},{reg:'231302023',name:'Rahul Kumar'},{reg:'231302024',name:'Himanshu Singh'},{reg:'231302025',name:'Samyak Jain'},{reg:'231302026',name:'Anshul'},{reg:'231302028',name:'Vikash Srivastava'},{reg:'231302031',name:'Varun Yadav'},{reg:'231302033',name:'Aman'},{reg:'231302034',name:'Anurag'},{reg:'231302035',name:'Sahil Singh'},{reg:'231302037',name:'Lavisha'},{reg:'231302038',name:'Umesh'},{reg:'231302045',name:'Aakash'},{reg:'231302047',name:'Arshad Khan'},{reg:'231302049',name:'Deepesh'},{reg:'231302051',name:'Kunal'},{reg:'231302054',name:'Harsh Chauhan'},{reg:'231302055',name:'Lakshita'},{reg:'231302058',name:'Pratishtha'},{reg:'231302061',name:'Sarthak'},{reg:'231302062',name:'Hunny Gulia'},{reg:'231302065',name:'Dushyant Vasisht'},{reg:'231302066',name:'Mayank Goyal'},{reg:'231302072',name:'Shubham Dagar'},{reg:'231302073',name:'Mahesh Kumar'},{reg:'231302075',name:'Shubham'},{reg:'231302081',name:'Tanush Banduni'},{reg:'231302082',name:'Anjali'},{reg:'231302083',name:'Himanshu Soni'},{reg:'231302085',name:'Shashank Verma'},{reg:'231302086',name:'Harsh'},{reg:'231302097',name:'Megha Sharma'},{reg:'231302099',name:'Mayank Garg'},{reg:'231302101',name:'Punit'},{reg:'231302102',name:'Deepanshu'},{reg:'231302105',name:'Tarun Yadav'},{reg:'231302107',name:'Priya'},{reg:'231302110',name:'Arpit'},{reg:'231302111',name:'Rohan'},{reg:'231302113',name:'Rajat Yadav'},{reg:'231302118',name:'Vandana Kumari'},{reg:'231302121',name:'Abhimanyu Kandhal'},{reg:'231302128',name:'Raghav Kumar Jha'},{reg:'231302133',name:'Rahul Kumar'},{reg:'231302136',name:'Prashant Rohilla'},{reg:'231302140',name:'Vanshika Rathi'},{reg:'231302142',name:'Nikhil Narang'},{reg:'231302144',name:'Yuvraj Singh Yadav'},{reg:'231302149',name:'Yash'},{reg:'231302151',name:'Pushpender Yadav'},{reg:'231302152',name:'Devesh Dwivedi'},{reg:'231302156',name:'Angel Suri'},{reg:'231302157',name:'Vishal Kumar Upadhyay'},{reg:'231302163',name:'Shubham Bhardwaj'},{reg:'231302166',name:'Manish Rathi'},{reg:'231302175',name:'Pinki Badhana'},{reg:'231302186',name:'Deepak Kumar'},{reg:'231302192',name:'Saloni'},{reg:'231302194',name:'Robin'},{reg:'231302203',name:'Varsha'},{reg:'231302205',name:'Sneha'},{reg:'231302214',name:'Monu'},{reg:'231302216',name:'Arpit Gupta'},{reg:'231302221',name:'Nikshey Yadav'},{reg:'231302230',name:'Aditya Chauhan'},{reg:'231302232',name:'Diksha Sharma'},{reg:'231302244',name:'Naman Sharma'},{reg:'231302257',name:'Lakshay Rohilla'}],
      'B': [{reg:'231302022',name:'Himanshu'},{reg:'231302032',name:'Aditya'},{reg:'231302043',name:'Deepak'},{reg:'231302050',name:'Savitender Singh'},{reg:'231302108',name:'Riya Yadav'},{reg:'231302123',name:'Yusuf Safi'},{reg:'231302148',name:'Pratik Dagar'},{reg:'231302160',name:'Piyush Kumar'},{reg:'231302171',name:'Harsh Yadav'},{reg:'231302173',name:'Depinder Kumar'},{reg:'231302178',name:'Jasmeet'},{reg:'231302188',name:'Aman Yadav'},{reg:'231302191',name:'Nitin Kumar'},{reg:'231302200',name:'Vinit Sheetal'},{reg:'231302201',name:'Anshu'},{reg:'231302204',name:'Samarth Ratwaya'},{reg:'231302209',name:'Ankit'},{reg:'231302215',name:'Vishal'},{reg:'231302218',name:'Vansh Rana'},{reg:'231302222',name:'Neeraj Lodhi'},{reg:'231302224',name:'Rudransh Bhardwaj'},{reg:'231302226',name:'Siddharth Sehwag'},{reg:'231302235',name:'Naman Kumar'},{reg:'231302237',name:'Nishant Choudhary'},{reg:'231302245',name:'Parikshit Gulia'},{reg:'231302253',name:'Trish Dalal'},{reg:'231302258',name:'Priyanshu'},{reg:'231302262',name:'Ashish Verma'},{reg:'231302263',name:'Ranu Raj'},{reg:'241302030',name:'Archi'},{reg:'241302267',name:'Jatin'},{reg:'241302279',name:'Vimarsh'},{reg:'221302122',name:'Mayank'},{reg:'231302001',name:'Yash Gupta'},{reg:'231302013',name:'Kishlay Kisu'},{reg:'231302018',name:'Abhishek Pal'},{reg:'231302080',name:'Aryan Sharma'},{reg:'231302094',name:'Rohit Kumar'},{reg:'231302127',name:'Shubranshu Sekher Sahoo'},{reg:'231302129',name:'Kunal Vashist'},{reg:'231302137',name:'Muskan'},{reg:'231302138',name:'Pooja'},{reg:'231302177',name:'Adarsh Kumar'},{reg:'231302190',name:'Lakshay Sharma'},{reg:'231302193',name:'Usha'},{reg:'231302211',name:'Aditya Tokas'},{reg:'231302213',name:'Priyanshu'},{reg:'231302217',name:'Mukul Punia'},{reg:'231302220',name:'Mishra Suyansh Naresh'},{reg:'231302233',name:'Ayush Gupta'},{reg:'231302252',name:'Rohan'},{reg:'241302251',name:'Rait Ranvirsingh Surinder Singh'},{reg:'241302050',name:'Rajan'}],
      'C': [{reg:'231302009',name:'Himanshu'},{reg:'231302012',name:'Prince Kumar'},{reg:'231302014',name:'Harish'},{reg:'231302016',name:'Arpit Sharma'},{reg:'231302017',name:'Tushar Singh'},{reg:'231302019',name:'Arjun Singh Dogra'},{reg:'231302021',name:'Ankit'},{reg:'231302027',name:'Janvi'},{reg:'231302029',name:'Himani Gupta'},{reg:'231302030',name:'Mohammad Ali'},{reg:'231302036',name:'Ishan'},{reg:'231302039',name:'Aryan Kumar'},{reg:'231302040',name:'Vishal'},{reg:'231302041',name:'Karan Sharma'},{reg:'231302042',name:'Anshuman Singh'},{reg:'231302044',name:'Gaurav Yadav'},{reg:'231302046',name:'Aakash Bisht'},{reg:'231302048',name:'Anvi Goyal'},{reg:'231302052',name:'Vishal'},{reg:'231302056',name:'Devanshi Sharma'},{reg:'231302057',name:'Neha Sharma'},{reg:'231302059',name:'Lekit Yadav'},{reg:'231302060',name:'Rahul Bajaj'},{reg:'231302063',name:'Riya'},{reg:'231302067',name:'Varun Raghav'},{reg:'231302068',name:'Faiza'},{reg:'231302069',name:'Deepak Yadav'},{reg:'231302070',name:'Nirupam Sharma'},{reg:'231302071',name:'Deepanshu Jangra'},{reg:'231302078',name:'Manjeet'},{reg:'231302079',name:'Jasneet Kaur'},{reg:'231302088',name:'Deepanshu Rai'},{reg:'231302089',name:'Sidhanshu'},{reg:'231302093',name:'Vaibhav'},{reg:'231302095',name:'Ayush Parashar'},{reg:'231302100',name:'Ritika Sharma'},{reg:'231302103',name:'Akash Rana'},{reg:'231302104',name:'Pallavi Mudgal'},{reg:'231302106',name:'Chetan Sagar'},{reg:'231302109',name:'Arvind Kumar'},{reg:'231302112',name:'Daksh Yadav'},{reg:'231302114',name:'Guneet Dhaka'},{reg:'231302115',name:'Aaditya Kaushik'},{reg:'231302116',name:'Jatin Nagarwal'},{reg:'231302117',name:'Sneha Boora'},{reg:'231302120',name:'Kushagar Siwach'},{reg:'231302122',name:'Renu'},{reg:'231302125',name:'Manish Dahiya'},{reg:'231302130',name:'Tanuj Bhardwaj'},{reg:'231302131',name:'Karan Singh Aswal'},{reg:'231302135',name:'Aneek Raj'},{reg:'231302139',name:'Rahul'},{reg:'231302141',name:'Vivek'},{reg:'231302143',name:'Jai Anmol Arora'},{reg:'231302150',name:'Noama Syed'},{reg:'231302153',name:'Dhairya Singh'},{reg:'231302154',name:'Goutam Verma'},{reg:'231302155',name:'Vanshika'},{reg:'231302158',name:'Nishant'},{reg:'231302159',name:'Krrish Malhotra'},{reg:'231302161',name:'Tanya Singh'},{reg:'231302162',name:'Rohit Korpal'}],
      'D': [{reg:'231302164',name:'Priyanshu'},{reg:'231302165',name:'Purav Tanwar'},{reg:'231302167',name:'Kislay Pandey'},{reg:'231302169',name:'Arpit Pandey'},{reg:'231302174',name:'Vishav Garg'},{reg:'231302176',name:'Sukhsham Sharma'},{reg:'231302181',name:'Devesh Yadav'},{reg:'231302182',name:'Aditya Vats'},{reg:'231302185',name:'Kunal Sharma'},{reg:'231302189',name:'Saurav'},{reg:'231302196',name:'Vrishank Sharma'},{reg:'231302199',name:'Parteek'},{reg:'231302206',name:'Dharam Singh'},{reg:'231302207',name:'Nikhil Kumar'},{reg:'231302208',name:'Punyam'},{reg:'231302210',name:'Pushpender Yadav'},{reg:'231302212',name:'Deepanshu'},{reg:'231302223',name:'Anurag'},{reg:'231302227',name:'Sandeep Singh'},{reg:'231302228',name:'Khushi Singhal'},{reg:'231302231',name:'Tanvir Singh Wasir'},{reg:'231302234',name:'Vanshika Bisht'},{reg:'231302239',name:'Aanchal'},{reg:'231302247',name:'Aakash Dagar'},{reg:'231302250',name:'Aakash'},{reg:'231302251',name:'Himanshu Aggarwal'},{reg:'231302255',name:'Anshika'},{reg:'231302256',name:'Tanishk'},{reg:'231302266',name:'Dhawan Sawant'},{reg:'241302032',name:'Arunava Chakraborty'},{reg:'231302004',name:'Vivek Singh'},{reg:'231302005',name:'Himanshi'},{reg:'231302008',name:'Aryan'},{reg:'231302010',name:'Hiten'},{reg:'231302011',name:'Muskan'},{reg:'231302015',name:'Devanshu Shekhar'},{reg:'231302084',name:'Nirmit'},{reg:'231302087',name:'Vikash Meena'},{reg:'231302091',name:'Muskan'},{reg:'231302092',name:'Yashika Janghu'},{reg:'231302124',name:'Harsh Vashisth'},{reg:'231302145',name:'Rajat Mallick'},{reg:'231302170',name:'Nikita Sengar'},{reg:'231302183',name:'Ayush Yadav'},{reg:'231302184',name:'Ayush Gulia'},{reg:'231302197',name:'Uday Pratap Singh'},{reg:'231302219',name:'Manish Kumar'},{reg:'231302229',name:'Prikshit'},{reg:'231302246',name:'Saket'},{reg:'231302248',name:'Sambhav'},{reg:'231302254',name:'Tushar Singh Tanwar'},{reg:'231302264',name:'Katrine Alice Daniel'}]
    },
    'MATLAB': {
      'C': [{reg:'241302002',name:'Rakshita Hooda'},{reg:'241302004',name:'Prarit Sharma'},{reg:'241302005',name:'Krrish Sharma'},{reg:'241302008',name:'Tarun'},{reg:'241302010',name:'Kiran'},{reg:'241302013',name:'Mayank Nehra'},{reg:'241302017',name:'Mohit Kumar Mishra'},{reg:'241302020',name:'Vanshik Sehrawat'},{reg:'241302021',name:'Suryansh Choudhary'},{reg:'241302029',name:'Jatin Chauhan'},{reg:'241302033',name:'Priyanshu'},{reg:'241302036',name:'Nikhil'},{reg:'241302039',name:'Kyan V Singh'},{reg:'241302042',name:'Vincent Chinlal'},{reg:'241302049',name:'Vansh Gulia'},{reg:'241302051',name:'Shreya Singh'},{reg:'241302052',name:'Ngachammi Khapudang'},{reg:'241302053',name:'Chaitanya Garg'},{reg:'241302055',name:'Umang Sharma'},{reg:'241302057',name:'Prem Kumar'},{reg:'241302058',name:'Gourav Shokeen'},{reg:'241302059',name:'Rachna Rai'},{reg:'241302065',name:'Shaili Yadav'},{reg:'241302068',name:'Sarika Yadav'},{reg:'241302069',name:'Sameer Yadav'},{reg:'241302070',name:'Ruben Rajeev George'},{reg:'241302071',name:'Divyansh Tanwar'},{reg:'241302075',name:'Rahul Gulia'},{reg:'241302077',name:'Sahil Choudhary'},{reg:'241302078',name:'Rahul'},{reg:'241302079',name:'Yash Aswal'},{reg:'241302081',name:'Harshit Khemani'},{reg:'241302087',name:'Sanket Kumar Prasad'},{reg:'241302090',name:'Omyansh Rawat'},{reg:'241302096',name:'Chalcy'},{reg:'241302097',name:'Yash Phalswal'},{reg:'241302099',name:'Manmeet Saroa'},{reg:'241302100',name:'Madhav Bassi'},{reg:'241302101',name:'Amit Kumar Patra'},{reg:'241302104',name:'Ekta'},{reg:'241302106',name:'Gaurav Joon'},{reg:'241302107',name:'Binit Kumar'},{reg:'241302108',name:'Tanisha'},{reg:'241302109',name:'Kushagra Agrawal'},{reg:'241302112',name:'Randeep Gir'},{reg:'241302113',name:'Harshit Singh'},{reg:'241302117',name:'Geetika'},{reg:'241302118',name:'Harshit'},{reg:'241302122',name:'Kush Ahuja'},{reg:'241302125',name:'Keshav Sood'},{reg:'241302126',name:'Rishabh'},{reg:'241302131',name:'Vaibhav Soni'},{reg:'241302133',name:'Vansh Sindhu'},{reg:'241302138',name:'Kanupriya'},{reg:'241302140',name:'Rijul Jain'},{reg:'241302141',name:'Eleesa Merin Jacob'},{reg:'241302142',name:'Joncy'}]
    }
  };

  if (section === 'ALL') return R[subject] || {};
  return (R[subject] && R[subject][section]) ? R[subject][section] : [];
}
