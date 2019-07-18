var SCRIPT_NAME = 'Utils'
var SCRIPT_VERSION = 'v1.1'

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