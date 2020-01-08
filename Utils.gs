var SCRIPT_NAME = 'Utils'
var SCRIPT_VERSION = 'v1.3'

var DateDiff = (function(ns) {

  // Get the number of whole days
  ns.inDays = function(d1, d2) {  
    checkParams(d1, d2)    
    return Math.floor((d2 - d1) / (24 * 3600 * 1000))
  }
  
  ns.inWeeks = function(d1, d2) {  
    checkParams(d1, d2)        
    return parseInt((d2 - d1)/(24 * 3600 * 1000 * 7));
  }
  
  ns.inMonths = function(d1, d2) {
  
    checkParams(d1, d2)    
    
    var d1Y = d1.getFullYear();
    var d2Y = d2.getFullYear();
    var d1M = d1.getMonth();
    var d2M = d2.getMonth();
    
    return (d2M + 12 * d2Y) - (d1M + 12 * d1Y);
  }
  
  ns.inYears = function(d1, d2) {
    checkParams(d1, d2)    
    return d2.getFullYear() - d1.getFullYear();
  }
  
  // Private Functions
  // -----------------
  
  function checkParams(d1, d2) {
    if (!(d1 instanceof Date) || !(d2 instanceof Date)) {
      throw new Error('DateDiff - bad args. d1: ' + d1 + ', d2:' + d2)
    }
  }
  
  return ns;
  
})(DateDiff || {})

function arrayDiff(arr1, arr2) {
  var newArr = arr1.concat(arr2)  
  return newArr.filter(function (item) { 
    if (arr1.indexOf(item) === -1 || arr2.indexOf(item) === -1) { 
      return item; 
    }
  })
}

/**
 * Get the columns numbers of the various named ranges from the submissions
 * from the Promotions Request Form
 *
 * @param {Spreadsheet}
 *
 * @return {object} {[column name] : [column number]}
 */

function getPRFColumns(spreadsheet) {

  if (!spreadsheet) {
    throw new Error('No spreadsheet arg')
  }

  var namedRanges = spreadsheet.getNamedRanges();
  var columnNumbers = {};
  
  for (var key in namedRanges) {  
    var namedRange = namedRanges[key];
    if (namedRange.getRange().getRow() === 2) {
      columnNumbers[namedRange.getName()] = namedRange.getRange().getColumn();
    }
  }
  
  return columnNumbers
  
} // getPRFColumns()

function getDoc(testDocId) {
  var doc = DocumentApp.getActiveDocument()
  if (doc === null) {
    if (testDocId) {
      doc = DocumentApp.openById(testDocId)
    } else {
      throw new Error('No test doc ID')
    }
  }
  return doc
}

function getUi() {
  var doc = DocumentApp.getActiveDocument()
  var ui = null
  if (doc !== null) {
    ui = DocumentApp.getUi();
  }
  return ui
}

/* To Title Case © 2018 David Gouch | https://github.com/gouch/to-title-case */

// eslint-disable-next-line no-extend-native
function toTitleCase(title) {
  'use strict'
  var smallWords = /^(a|an|and|as|at|but|by|en|for|if|in|nor|of|on|or|per|the|to|v.?|vs.?|via)$/i
  var alphanumericPattern = /([A-Za-z0-9\u00C0-\u00FF])/
  var wordSeparators = /([ :–—-])/

  return title.split(wordSeparators)
    .map(function (current, index, array) {
      if (
        /* Check for small words */
        current.search(smallWords) > -1 &&
        /* Skip first and last word */
        index !== 0 &&
        index !== array.length - 1 &&
        /* Ignore title end and subtitle start */
        array[index - 3] !== ':' &&
        array[index + 1] !== ':' &&
        /* Ignore small words that start a hyphenated phrase */
        (array[index + 1] !== '-' ||
          (array[index - 1] === '-' && array[index + 1] === '-'))
      ) {
        return current.toLowerCase()
      }

      /* Ignore intentional capitalization */
      if (current.substr(1).search(/[A-Z]|\../) > -1) {
        return current
      }

      /* Ignore URLs */
      if (array[index + 1] === ':' && array[index + 2] !== '') {
        return current
      }

      /* Capitalize the first letter */
      return current.replace(alphanumericPattern, function (match) {
        return match.toUpperCase()
      })
    })
    .join('')
}

function toSentenceCase(sentence) {
  if (sentence === undefined || sentence === '') {
    return '';
  }
  var toSentenceCase = sentence[0].toUpperCase() + sentence.slice(1);
  if (sentence[sentence.length - 1] !== '.') {
    toSentenceCase += '.';
  }
  return toSentenceCase;
}

function fDate(date, format){//returns the date formatted with format, default to today if date not provided
  date = date || new Date()
  format = format || "MM/dd/yy"
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), format)
}

function getFormatedDateForEvent(date){
  //Log_.fine( '--getFormatedDateForEvent('+fDate(date)+')' )
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE, MMMM d 'at' h:mma")
  //for compatability with old script output, lowercase the meridiem
  return formattedDate.replace(/[A,P]M$/, function(l){ return l.toLowerCase() })
  //or just //return formattedDate.replace('PM', 'pm').replace('AM', 'am');
}

function getUpcomingSunday_TEST() {
  //these should all be the same date (time will be diff on the default)
  
  var date = new Date()
  date.setHours(23,0,0,0)
  //Log_.fine(date)
  //Log_.fine( fDate( getUpcomingSunday(date) ) )
  
}
function getUpcomingSunday(date, skipTodayIfSunday) {
  //return the next Sunday, which might be today
  //skipTodayIfSunday skips this Sunday and returns next week Sunday
  //Log_.fine( '--getUpcomingSunday_('+(date ? fDate(date) : 'null')+')' )
  date = new Date(date || new Date());//clone the date so as not to change the original
  date.setHours(0,0,0,0)
  if( skipTodayIfSunday || date.getDay() >0)//if it's not a Sunday...
    date.setDate(date.getDate() -date.getDay() +7);//subtract days to get to Sunday then add a week
  //Log_.fine('upcomingSunday returned: '+fDate(date))
  return date
}

function searchInMaster(masterId, str){//find first occurence of str
  //Log_.info('--searchInMaster_('+str+')')
  str = str.replace(/\[/, '\\[').replace(/\./, '\\.')
  //really should replace all regex reserved chars but we only need find like "[ 04.22 ] Sunday Announcements"
  var body = getMasterBody(masterId)//  var body = DocumentApp.openById(id).getBody()
  
  var hit = body.findText(str)
  
  //Log_.fine('hit:' +hit)
  return hit ? body.getChildIndex( hit.getElement().getParent()) : null

}

function getMasterBody(masterId){
  //usage: var body = getMasterBody_();//  var body = DocumentApp.openById(id).getBody()
  try{
    //var masterDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID'))
    var masterDoc = DocumentApp.openById(masterId)
  }catch(e){
    throw new Error('Unable to open Master Announcements file.  Check permissions and be sure it has been set in CONFIG_.files.masterAnnouncementsID')
  }
  return masterDoc.getBody()
}

function dates_TEST() {
  var d1 = new Date(2018,0,10,0)
  var d2 = new Date(2018,0,11,1)
  var d = DateDiff.inDays(d1, d2)
  //Log_.fine(d)
}

/**
* Add time to a date in specified interval
* Negative values work as well
*
* @param {date} javascript datetime object
* @param {interval} text interval name [year|quarter|month|week|day|hour|minute|second]
* @param {units} integer units of interval to add to date
* @return {date object} 
*/
function dateAdd(date, interval, units) {
  date = new Date(date); //don't change original date
  switch(interval.toLowerCase()) {
    case 'year'   :  date.setFullYear(date.getFullYear() + units);            break;
    case 'quarter':  date.setMonth   (date.getMonth()    + units*3);          break;
    case 'month'  :  date.setMonth   (date.getMonth()    + units);            break;
    case 'week'   :  date.setDate    (date.getDate()     + units*7);          break;
    case 'day'    :  date.setDate    (date.getDate()     + units);            break;
    case 'hour'   :  date.setTime    (date.getTime()     + units*60*60*1000); break;
    case 'minute' :  date.setTime    (date.getTime()     + units*60*1000);    break;
    case 'second' :  date.setTime    (date.getTime()     + units*1000);       break;
    default       :  date = undefined; break;
  }
  return date
}