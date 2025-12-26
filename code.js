/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * 🌙✨ MAGICKAL MORTALZ COVEN - Google Apps Script Web Application ✨🌙
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * 🔮 Server-Side Functions (Code.gs)
 * 📅 Created: 2025
 * 👨‍💻 Author: Erik Gaton (Traikdude) - Founder & High Priest
 * 🎨 UI Protocol: Joyful UI v2.1
 * 
 * 📂 File Structure:
 *   - Code.gs (this file) - Server-side functions
 *   - index.html - Main HTML structure
 *   - css.html - Stylesheet (Joyful UI)
 *   - js.html - Client-side JavaScript
 * 
 * 📊 Google Sheets Structure (Auto-Created):
 *   - Members - User accounts and profiles
 *   - CurriculumProgress - Year/degree tracking
 *   - Attendance - Sabbat/Esbat participation
 *   - Grimoire - Personal spell/ritual entries
 *   - ActivityLog - System activity tracking
 * 
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// 🚀 WEB APP ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 🌐 doGet() - Serves the web application
 * This is called when the web app URL is accessed
 */
function doGet(e) {
  Logger.log('🌙 Magickal Mortalz Coven App accessed');
  
  // 📋 Initialize sheets if they don't exist
  initializeSheets();
  
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('🌙 Magickal Mortalz Coven')
    .setFaviconUrl('https://em-content.zobj.net/source/apple/391/crescent-moon_1f319.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 📄 include() - Includes HTML partials (css.html, js.html)
 * @param {string} filename - The file to include
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ═══════════════════════════════════════════════════════════════════════════════
// 🗄️ SHEET INITIALIZATION & CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 📊 Sheet Configuration Object
 * Defines all sheets and their column headers
 */
const SHEET_CONFIG = {
  MEMBERS: {
    name: 'Members',
    headers: ['ID', 'Email', 'CraftName', 'RealName', 'JoinDate', 'CurrentDegree', 'Avatar', 'Bio', 'IsActive', 'LastLogin']
  },
  PROGRESS: {
    name: 'CurriculumProgress',
    headers: ['ID', 'MemberID', 'Year', 'Module', 'Status', 'StartDate', 'CompletedDate', 'Notes', 'InstructorApproval']
  },
  ATTENDANCE: {
    name: 'Attendance',
    headers: ['ID', 'MemberID', 'EventType', 'EventName', 'EventDate', 'Attended', 'Notes', 'RecordedBy']
  },
  GRIMOIRE: {
    name: 'Grimoire',
    headers: ['ID', 'MemberID', 'EntryType', 'Title', 'Content', 'Tags', 'CreatedDate', 'ModifiedDate', 'IsPrivate', 'Category']
  },
  ACTIVITY: {
    name: 'ActivityLog',
    headers: ['Timestamp', 'MemberID', 'Action', 'Details', 'IPAddress']
  }
};

/**
 * 🏗️ initializeSheets() - Creates all required sheets if they don't exist
 */
function initializeSheets() {
  const ss = getSpreadsheet();
  
  Object.values(SHEET_CONFIG).forEach(config => {
    let sheet = ss.getSheetByName(config.name);
    
    if (!sheet) {
      Logger.log(`📋 Creating sheet: ${config.name}`);
      sheet = ss.insertSheet(config.name);
      
      // 🎨 Set up headers with formatting
      const headerRange = sheet.getRange(1, 1, 1, config.headers.length);
      headerRange.setValues([config.headers]);
      headerRange.setBackground('#FF6B9D');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 📏 Auto-resize columns
      config.headers.forEach((_, i) => {
        sheet.autoResizeColumn(i + 1);
      });
      
      // ❄️ Freeze header row
      sheet.setFrozenRows(1);
    }
  });
  
  Logger.log('✅ All sheets initialized');
}

/**
 * 📁 getSpreadsheet() - Gets or creates the data spreadsheet
 * @returns {Spreadsheet} The data spreadsheet
 */
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  
  if (!ssId) {
    // 🆕 Create new spreadsheet
    const ss = SpreadsheetApp.create('🌙 Magickal Mortalz Coven Data');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
    Logger.log(`📊 Created new spreadsheet: ${ssId}`);
    
    // 🗑️ Delete default "Sheet1"
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet) {
      ss.deleteSheet(defaultSheet);
    }
  }
  
  return SpreadsheetApp.openById(ssId);
}

/**
 * 📋 getSheet() - Gets a specific sheet by config key
 * @param {string} configKey - Key from SHEET_CONFIG
 * @returns {Sheet} The requested sheet
 */
function getSheet(configKey) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(SHEET_CONFIG[configKey].name);
}

// ═══════════════════════════════════════════════════════════════════════════════
// 🔧 UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 🆔 generateId() - Generates a unique ID
 * @param {string} prefix - Prefix for the ID (e.g., 'MEM', 'GRM')
 * @returns {string} Unique ID
 */
function generateId(prefix = 'ID') {
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${timestamp}-${random}`;
}

/**
 * 📅 formatDate() - Formats a date for display
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(date = new Date()) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 🔍 findRowByColumn() - Finds a row by matching a column value
 * @param {Sheet} sheet - The sheet to search
 * @param {number} column - Column index (1-based)
 * @param {*} value - Value to find
 * @returns {number} Row number or -1 if not found
 */
function findRowByColumn(sheet, column, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][column - 1] === value) {
      return i + 1;
    }
  }
  return -1;
}

/**
 * 📊 sheetToObjects() - Converts sheet data to array of objects
 * @param {Sheet} sheet - The sheet to convert
 * @returns {Array} Array of row objects
 */
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      let value = row[i];
      // 🛡️ Convert Date objects to ISO strings for serialization
      if (value instanceof Date) {
        value = value.toISOString();
      }
      obj[header] = value;
    });
    return obj;
  });
}

/**
 * 📝 logActivity() - Logs user activity
 * @param {string} memberId - Member's ID
 * @param {string} action - Action performed
 * @param {string} details - Additional details
 */
function logActivity(memberId, action, details) {
  try {
    const sheet = getSheet('ACTIVITY');
    sheet.appendRow([
      formatDate(),
      memberId || 'SYSTEM',
      action,
      details,
      ''  // IP placeholder
    ]);
  } catch (e) {
    Logger.log(`⚠️ Activity log error: ${e.message}`);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// 👤 MEMBER MANAGEMENT (CRUD)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 🆕 createMember() - Creates a new member
 * @param {Object} memberData - Member information
 * @returns {Object} Result with success status and member ID
 */
function createMember(memberData) {
  try {
    Logger.log(`🆕 Creating member: ${JSON.stringify(memberData)}`);
    
    const sheet = getSheet('MEMBERS');
    const id = generateId('MEM');
    
    const row = [
      id,
      memberData.email || '',
      memberData.craftName || '',
      memberData.realName || '',
      formatDate(),
      'Neophyte',  // 🌱 Everyone starts as Neophyte
      memberData.avatar || '🧙‍♂️',
      memberData.bio || '',
      true,
      formatDate()
    ];
    
    sheet.appendRow(row);
    logActivity(id, 'MEMBER_CREATED', `New member: ${memberData.craftName}`);
    
    // 📚 Initialize curriculum progress for Year 1
    initializeCurriculumForMember(id);
    
    Logger.log(`✅ Member created: ${id}`);
    
    // 🛡️ Return clean serializable object
    const result = { success: true, memberId: id, message: '🌙 Welcome to the Coven!' };
    return JSON.parse(JSON.stringify(result));
    
  } catch (e) {
    Logger.log(`❌ Error creating member: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

/**
 * 📖 getMember() - Gets a member by ID
 * @param {string} memberId - Member's ID
 * @returns {Object} Member data or null
 */
function getMember(memberId) {
  try {
    const sheet = getSheet('MEMBERS');
    const row = findRowByColumn(sheet, 1, memberId);
    
    if (row === -1) return null;
    
    const data = sheet.getRange(row, 1, 1, SHEET_CONFIG.MEMBERS.headers.length).getValues()[0];
    const member = {};
    SHEET_CONFIG.MEMBERS.headers.forEach((header, i) => {
      let value = data[i];
      // 🛡️ Convert Date objects to ISO strings for serialization
      if (value instanceof Date) {
        value = value.toISOString();
      }
      member[header] = value;
    });
    
    return member;
    
  } catch (e) {
    Logger.log(`❌ Error getting member: ${e.message}`);
    return null;
  }
}

/**
 * 📋 getAllMembers() - Gets all active members
 * @returns {Array} Array of member objects
 */
function getAllMembers() {
  try {
    const sheet = getSheet('MEMBERS');
    const members = sheetToObjects(sheet);
    return members.filter(m => m.IsActive === true || m.IsActive === 'TRUE');
  } catch (e) {
    Logger.log(`❌ Error getting members: ${e.message}`);
    return [];
  }
}

/**
 * ✏️ updateMember() - Updates a member's information
 * @param {string} memberId - Member's ID
 * @param {Object} updates - Fields to update
 * @returns {Object} Result with success status
 */
function updateMember(memberId, updates) {
  try {
    const sheet = getSheet('MEMBERS');
    const row = findRowByColumn(sheet, 1, memberId);
    
    if (row === -1) {
      return { success: false, error: 'Member not found' };
    }
    
    const headers = SHEET_CONFIG.MEMBERS.headers;
    
    Object.keys(updates).forEach(key => {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(row, colIndex + 1).setValue(updates[key]);
      }
    });
    
    logActivity(memberId, 'MEMBER_UPDATED', JSON.stringify(updates));
    return { success: true, message: '✨ Profile updated!' };
    
  } catch (e) {
    Logger.log(`❌ Error updating member: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// 📚 CURRICULUM PROGRESS TRACKING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 🎓 Curriculum Structure
 * Defines the 4-year initiatory path
 */
const CURRICULUM = {
  year1: {
    name: 'Neophyte',
    emoji: '🌱',
    modules: [
      'Meditation Basics',
      'History of Wicca',
      'Tool Consecration',
      'The Wheel of the Year',
      'Basic Energy Work',
      'Circle Casting Fundamentals',
      'Introduction to Deities',
      'Ethics & The Rede'
    ]
  },
  year2: {
    name: '1st Degree Initiate',
    emoji: '🛡️',
    modules: [
      'Protection & Psychic Defense',
      'Advanced Circle Work',
      'Elemental Invocations',
      'Herbology Basics',
      'Crystal Correspondences',
      'Moon Phases & Esbats',
      'Sabbat Celebrations',
      'Personal Grimoire Creation'
    ]
  },
  year3: {
    name: '2nd Degree Initiate',
    emoji: '🌘',
    modules: [
      'Goddess Mysteries',
      'Shadow Work',
      'Advanced Divination',
      'Astral Projection',
      'Trance & Journey Work',
      'Ritual Leadership',
      'Teaching Fundamentals',
      'Coven Dynamics'
    ]
  },
  year4: {
    name: '3rd Degree Initiate',
    emoji: '☀️',
    modules: [
      'God Mysteries',
      'Solar Magick',
      'High Priesthood Studies',
      'Coven Leadership',
      'Initiation Facilitation',
      'Advanced Spellcraft',
      'Community Building',
      'Tradition Mastery'
    ]
  }
};

/**
 * 📚 initializeCurriculumForMember() - Sets up Year 1 modules for new member
 * @param {string} memberId - Member's ID
 */
function initializeCurriculumForMember(memberId) {
  try {
    const sheet = getSheet('PROGRESS');
    
    CURRICULUM.year1.modules.forEach((module, index) => {
      const id = generateId('PRG');
      sheet.appendRow([
        id,
        memberId,
        1,  // Year 1
        module,
        index === 0 ? 'In Progress' : 'Not Started',
        index === 0 ? formatDate() : '',
        '',
        '',
        ''
      ]);
    });
    
    Logger.log(`📚 Curriculum initialized for member: ${memberId}`);
    
  } catch (e) {
    Logger.log(`❌ Error initializing curriculum: ${e.message}`);
  }
}

/**
 * 📊 getMemberProgress() - Gets curriculum progress for a member
 * @param {string} memberId - Member's ID
 * @returns {Object} Progress data organized by year
 */
function getMemberProgress(memberId) {
  try {
    const sheet = getSheet('PROGRESS');
    const allProgress = sheetToObjects(sheet);
    const memberProgress = allProgress.filter(p => p.MemberID === memberId);
    
    // 📈 Calculate statistics
    const stats = {
      totalModules: memberProgress.length,
      completed: memberProgress.filter(p => p.Status === 'Completed').length,
      inProgress: memberProgress.filter(p => p.Status === 'In Progress').length,
      notStarted: memberProgress.filter(p => p.Status === 'Not Started').length
    };
    
    stats.percentComplete = stats.totalModules > 0 
      ? Math.round((stats.completed / stats.totalModules) * 100) 
      : 0;
    
    // 📋 Organize by year
    const byYear = {};
    [1, 2, 3, 4].forEach(year => {
      byYear[`year${year}`] = memberProgress.filter(p => p.Year == year);
    });
    
    return {
      stats,
      byYear,
      curriculum: CURRICULUM
    };
    
  } catch (e) {
    Logger.log(`❌ Error getting progress: ${e.message}`);
    return { stats: {}, byYear: {}, curriculum: CURRICULUM };
  }
}

/**
 * ✅ updateModuleStatus() - Updates a module's completion status
 * @param {string} progressId - Progress record ID
 * @param {string} status - New status ('Not Started', 'In Progress', 'Completed')
 * @returns {Object} Result with success status
 */
function updateModuleStatus(progressId, status) {
  try {
    const sheet = getSheet('PROGRESS');
    const row = findRowByColumn(sheet, 1, progressId);
    
    if (row === -1) {
      return { success: false, error: 'Progress record not found' };
    }
    
    // 📝 Update status
    sheet.getRange(row, 5).setValue(status);
    
    // 📅 Update dates based on status
    if (status === 'In Progress') {
      sheet.getRange(row, 6).setValue(formatDate());
    } else if (status === 'Completed') {
      sheet.getRange(row, 7).setValue(formatDate());
    }
    
    const module = sheet.getRange(row, 4).getValue();
    const memberId = sheet.getRange(row, 2).getValue();
    logActivity(memberId, 'MODULE_STATUS_UPDATED', `${module}: ${status}`);
    
    return { success: true, message: `✨ Module marked as ${status}!` };
    
  } catch (e) {
    Logger.log(`❌ Error updating module: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * 🎓 advanceMemberToNextYear() - Advances member to next curriculum year
 * @param {string} memberId - Member's ID
 * @returns {Object} Result with success status
 */
function advanceMemberToNextYear(memberId) {
  try {
    const member = getMember(memberId);
    if (!member) {
      return { success: false, error: 'Member not found' };
    }
    
    // 🎯 Determine current and next year
    const degreeMap = {
      'Neophyte': { current: 1, next: 2, nextDegree: '1st Degree Initiate' },
      '1st Degree Initiate': { current: 2, next: 3, nextDegree: '2nd Degree Initiate' },
      '2nd Degree Initiate': { current: 3, next: 4, nextDegree: '3rd Degree Initiate' },
      '3rd Degree Initiate': { current: 4, next: null, nextDegree: 'High Priest/ess' }
    };
    
    const progression = degreeMap[member.CurrentDegree];
    if (!progression || !progression.next) {
      return { success: false, error: 'Already at highest degree!' };
    }
    
    // 📚 Add next year's modules
    const sheet = getSheet('PROGRESS');
    const yearKey = `year${progression.next}`;
    
    CURRICULUM[yearKey].modules.forEach((module, index) => {
      const id = generateId('PRG');
      sheet.appendRow([
        id,
        memberId,
        progression.next,
        module,
        index === 0 ? 'In Progress' : 'Not Started',
        index === 0 ? formatDate() : '',
        '',
        '',
        ''
      ]);
    });
    
    // ✨ Update member degree
    updateMember(memberId, { CurrentDegree: progression.nextDegree });
    logActivity(memberId, 'DEGREE_ADVANCED', `Advanced to ${progression.nextDegree}`);
    
    return { 
      success: true, 
      message: `🎉 Congratulations! Advanced to ${progression.nextDegree}!`,
      newDegree: progression.nextDegree
    };
    
  } catch (e) {
    Logger.log(`❌ Error advancing member: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// 📅 ATTENDANCE TRACKING (SABBATS & ESBATS)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 🎃 Sabbat Calendar
 * The 8 Sabbats of the Wheel of the Year
 */
const SABBATS = [
  { name: 'Samhain', date: '10-31', emoji: '🎃', description: 'Witches\' New Year' },
  { name: 'Yule', date: '12-21', emoji: '🎄', description: 'Winter Solstice' },
  { name: 'Imbolc', date: '02-01', emoji: '🕯️', description: 'First Light of Spring' },
  { name: 'Ostara', date: '03-20', emoji: '🐣', description: 'Spring Equinox' },
  { name: 'Beltane', date: '05-01', emoji: '🔥', description: 'Fire Festival' },
  { name: 'Litha', date: '06-21', emoji: '☀️', description: 'Summer Solstice' },
  { name: 'Lughnasadh', date: '08-01', emoji: '🌾', description: 'First Harvest' },
  { name: 'Mabon', date: '09-22', emoji: '🍂', description: 'Autumn Equinox' }
];

/**
 * 📅 recordAttendance() - Records attendance for an event
 * @param {string} memberId - Member's ID
 * @param {Object} eventData - Event information
 * @returns {Object} Result with success status
 */
function recordAttendance(memberId, eventData) {
  try {
    const sheet = getSheet('ATTENDANCE');
    const id = generateId('ATT');
    
    sheet.appendRow([
      id,
      memberId,
      eventData.eventType || 'Sabbat',
      eventData.eventName,
      eventData.eventDate || formatDate(),
      eventData.attended !== false,
      eventData.notes || '',
      eventData.recordedBy || 'SELF'
    ]);
    
    logActivity(memberId, 'ATTENDANCE_RECORDED', `${eventData.eventName}: ${eventData.attended ? 'Present' : 'Absent'}`);
    
    return { success: true, message: '✨ Attendance recorded!' };
    
  } catch (e) {
    Logger.log(`❌ Error recording attendance: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * 📊 getMemberAttendance() - Gets attendance history for a member
 * @param {string} memberId - Member's ID
 * @returns {Object} Attendance data with statistics
 */
function getMemberAttendance(memberId) {
  try {
    const sheet = getSheet('ATTENDANCE');
    const allAttendance = sheetToObjects(sheet);
    const memberAttendance = allAttendance.filter(a => a.MemberID === memberId);
    
    // 📈 Calculate statistics
    const sabbatAttendance = memberAttendance.filter(a => a.EventType === 'Sabbat');
    const esbatAttendance = memberAttendance.filter(a => a.EventType === 'Esbat');
    
    const stats = {
      totalEvents: memberAttendance.length,
      sabbatsAttended: sabbatAttendance.filter(a => a.Attended === true || a.Attended === 'TRUE').length,
      esbatsAttended: esbatAttendance.filter(a => a.Attended === true || a.Attended === 'TRUE').length,
      totalSabbats: sabbatAttendance.length,
      totalEsbats: esbatAttendance.length
    };
    
    return {
      records: memberAttendance,
      stats,
      sabbats: SABBATS
    };
    
  } catch (e) {
    Logger.log(`❌ Error getting attendance: ${e.message}`);
    return { records: [], stats: {}, sabbats: SABBATS };
  }
}

/**
 * 🌙 getUpcomingSabbats() - Gets upcoming sabbats for the next 3 months
 * @returns {Array} Array of upcoming sabbat objects
 */
function getUpcomingSabbats() {
  const today = new Date();
  const threeMonthsLater = new Date(today);
  threeMonthsLater.setMonth(threeMonthsLater.getMonth() + 3);
  
  const currentYear = today.getFullYear();
  const upcoming = [];
  
  SABBATS.forEach(sabbat => {
    const [month, day] = sabbat.date.split('-').map(Number);
    
    // Check this year and next year
    [currentYear, currentYear + 1].forEach(year => {
      const sabbatDate = new Date(year, month - 1, day);
      
      if (sabbatDate >= today && sabbatDate <= threeMonthsLater) {
        upcoming.push({
          name: sabbat.name,
          date: sabbat.date,
          emoji: sabbat.emoji,
          // 🛡️ Convert Date to ISO string for serialization
          fullDate: sabbatDate.toISOString(),
          dateString: Utilities.formatDate(sabbatDate, Session.getScriptTimeZone(), 'MMMM d, yyyy'),
          daysUntil: Math.ceil((sabbatDate - today) / (1000 * 60 * 60 * 24))
        });
      }
    });
  });
  
  // Sort by date string instead of Date object
  return upcoming.sort((a, b) => new Date(a.fullDate) - new Date(b.fullDate));
}

// ═══════════════════════════════════════════════════════════════════════════════
// 📜 GRIMOIRE (PERSONAL BOOK OF SHADOWS)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 📝 Grimoire Entry Categories
 */
const GRIMOIRE_CATEGORIES = [
  { id: 'spells', name: 'Spells & Workings', emoji: '🔮' },
  { id: 'rituals', name: 'Rituals', emoji: '⭕' },
  { id: 'divination', name: 'Divination Records', emoji: '🎴' },
  { id: 'dreams', name: 'Dream Journal', emoji: '💭' },
  { id: 'herbs', name: 'Herb Notes', emoji: '🌿' },
  { id: 'crystals', name: 'Crystal Work', emoji: '💎' },
  { id: 'deities', name: 'Deity Work', emoji: '🌙' },
  { id: 'meditation', name: 'Meditation Logs', emoji: '🧘' },
  { id: 'moon', name: 'Moon Workings', emoji: '🌕' },
  { id: 'other', name: 'Other Notes', emoji: '📝' }
];

/**
 * 📜 createGrimoireEntry() - Creates a new grimoire entry
 * @param {string} memberId - Member's ID
 * @param {Object} entryData - Entry information
 * @returns {Object} Result with success status and entry ID
 */
function createGrimoireEntry(memberId, entryData) {
  try {
    const sheet = getSheet('GRIMOIRE');
    const id = generateId('GRM');
    
    sheet.appendRow([
      id,
      memberId,
      entryData.entryType || 'Note',
      entryData.title || 'Untitled Entry',
      entryData.content || '',
      entryData.tags || '',
      formatDate(),
      formatDate(),
      entryData.isPrivate !== false,
      entryData.category || 'other'
    ]);
    
    logActivity(memberId, 'GRIMOIRE_ENTRY_CREATED', entryData.title);
    
    return { success: true, entryId: id, message: '📜 Entry added to your Grimoire!' };
    
  } catch (e) {
    Logger.log(`❌ Error creating grimoire entry: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * 📚 getGrimoireEntries() - Gets all grimoire entries for a member
 * @param {string} memberId - Member's ID
 * @param {string} category - Optional category filter
 * @returns {Object} Grimoire data with entries and statistics
 */
function getGrimoireEntries(memberId, category = null) {
  try {
    const sheet = getSheet('GRIMOIRE');
    const allEntries = sheetToObjects(sheet);
    let memberEntries = allEntries.filter(e => e.MemberID === memberId);
    
    // 🔍 Filter by category if specified
    if (category && category !== 'all') {
      memberEntries = memberEntries.filter(e => e.Category === category);
    }
    
    // 📊 Sort by most recent
    memberEntries.sort((a, b) => new Date(b.ModifiedDate) - new Date(a.ModifiedDate));
    
    // 📈 Calculate statistics
    const stats = {
      totalEntries: memberEntries.length,
      byCategory: {}
    };
    
    GRIMOIRE_CATEGORIES.forEach(cat => {
      stats.byCategory[cat.id] = allEntries.filter(
        e => e.MemberID === memberId && e.Category === cat.id
      ).length;
    });
    
    return {
      entries: memberEntries,
      stats,
      categories: GRIMOIRE_CATEGORIES
    };
    
  } catch (e) {
    Logger.log(`❌ Error getting grimoire entries: ${e.message}`);
    return { entries: [], stats: {}, categories: GRIMOIRE_CATEGORIES };
  }
}

/**
 * ✏️ updateGrimoireEntry() - Updates a grimoire entry
 * @param {string} entryId - Entry ID
 * @param {Object} updates - Fields to update
 * @returns {Object} Result with success status
 */
function updateGrimoireEntry(entryId, updates) {
  try {
    const sheet = getSheet('GRIMOIRE');
    const row = findRowByColumn(sheet, 1, entryId);
    
    if (row === -1) {
      return { success: false, error: 'Entry not found' };
    }
    
    const headers = SHEET_CONFIG.GRIMOIRE.headers;
    
    Object.keys(updates).forEach(key => {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(row, colIndex + 1).setValue(updates[key]);
      }
    });
    
    // 📅 Update modified date
    sheet.getRange(row, headers.indexOf('ModifiedDate') + 1).setValue(formatDate());
    
    const memberId = sheet.getRange(row, 2).getValue();
    logActivity(memberId, 'GRIMOIRE_ENTRY_UPDATED', updates.Title || entryId);
    
    return { success: true, message: '✨ Entry updated!' };
    
  } catch (e) {
    Logger.log(`❌ Error updating grimoire entry: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * 🗑️ deleteGrimoireEntry() - Deletes a grimoire entry
 * @param {string} entryId - Entry ID
 * @returns {Object} Result with success status
 */
function deleteGrimoireEntry(entryId) {
  try {
    const sheet = getSheet('GRIMOIRE');
    const row = findRowByColumn(sheet, 1, entryId);
    
    if (row === -1) {
      return { success: false, error: 'Entry not found' };
    }
    
    const memberId = sheet.getRange(row, 2).getValue();
    const title = sheet.getRange(row, 4).getValue();
    
    sheet.deleteRow(row);
    logActivity(memberId, 'GRIMOIRE_ENTRY_DELETED', title);
    
    return { success: true, message: '🗑️ Entry deleted!' };
    
  } catch (e) {
    Logger.log(`❌ Error deleting grimoire entry: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// 🌙 DASHBOARD DATA AGGREGATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 📊 getDashboardData() - Gets all data needed for the dashboard
 * @param {string} memberId - Member's ID
 * @returns {Object} Comprehensive dashboard data
 */
function getDashboardData(memberId) {
  try {
    Logger.log(`📊 Getting dashboard data for member: ${memberId}`);
    
    const member = getMember(memberId);
    const progress = getMemberProgress(memberId);
    const attendance = getMemberAttendance(memberId);
    const grimoire = getGrimoireEntries(memberId);
    const upcomingSabbats = getUpcomingSabbats();
    
    const result = {
      success: true,
      member: member,
      progress: progress,
      attendance: attendance,
      grimoire: grimoire,
      upcomingSabbats: upcomingSabbats,
      curriculum: CURRICULUM,
      sabbats: SABBATS,
      grimoireCategories: GRIMOIRE_CATEGORIES
    };
    
    // 🛡️ Force JSON serialization to catch any non-serializable data
    const serialized = JSON.parse(JSON.stringify(result));
    
    Logger.log(`✅ Dashboard data prepared successfully`);
    return serialized;
    
  } catch (e) {
    Logger.log(`❌ Error getting dashboard data: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

/**
 * 🧪 testDataCreation() - Creates test data for development
 * Run this function manually to populate test data
 */
function testDataCreation() {
  // 🆕 Create test member
  const result = createMember({
    email: 'test@magickalmortalz.com',
    craftName: 'Moonstone',
    realName: 'Test User',
    bio: 'A seeker of ancient wisdom'
  });
  
  if (result.success) {
    const memberId = result.memberId;
    
    // 📜 Add test grimoire entries
    createGrimoireEntry(memberId, {
      title: 'Full Moon Protection Spell',
      content: 'Cast under the light of the full moon...',
      category: 'spells',
      tags: 'protection, moon, banishing'
    });
    
    createGrimoireEntry(memberId, {
      title: 'Amethyst Dream Work',
      content: 'Place amethyst under pillow for prophetic dreams...',
      category: 'crystals',
      tags: 'amethyst, dreams, divination'
    });
    
    // 📅 Add test attendance
    recordAttendance(memberId, {
      eventType: 'Sabbat',
      eventName: 'Samhain 2024',
      eventDate: '2024-10-31',
      attended: true,
      notes: 'Honored the ancestors'
    });
    
    Logger.log(`✅ Test data created for member: ${memberId}`);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// 🔧 MENU & INITIALIZATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * 📋 onOpen() - Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🌙 Magickal Mortalz')
    .addItem('📊 Initialize Sheets', 'initializeSheets')
    .addItem('🧪 Create Test Data', 'testDataCreation')
    .addSeparator()
    .addItem('🌐 Open Web App', 'openWebApp')
    .addToUi();
}

/**
 * 🌐 openWebApp() - Opens the deployed web app URL
 */
function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Web App...');
}

// ═══════════════════════════════════════════════════════════════════════════════
// 🌙 END OF CODE.GS
// ═══════════════════════════════════════════════════════════════════════════════