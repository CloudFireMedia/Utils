var SCRIPT_NAME = 'Utils'
var SCRIPT_VERSION = 'v1.2'

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

function getDoc() {
  var doc = DocumentApp.getActiveDocument()
  if (doc === null) {
    doc = DocumentApp.openById(TEST_DOC_ID_)
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
  if (sentence[sentence.length] !== '.') {
    toSentenceCase += '.';
  }
  return toSentenceCase;
}

