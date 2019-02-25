var winax = require('..');
var interval = setInterval(function () {
	winax.peekAndDispatchMessages(); // allows ActiveX event to be dispatched
}, 50);

var application = new ActiveXObject('Excel.Application');

var connectionPoints = winax.getConnectionPoints(application);
var connectionPoint = connectionPoints[0];

// Excel Application events: https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events
connectionPoint.advise({
  WindowActivate: function (workbook, win) {
    console.log('WindowActivate event triggerred');
  },
  WindowDeactivate: function (workbook, win) {
    console.log('WindowDeactivate event triggerred');
  },
  WindowResize: function (workbook, win) {
    console.log('WindowResize event triggerred');
  },
  NewWorkbook: function (workbook) {
    console.log('NewWorkbook event triggerred');
    var worksheet = workbook.Worksheets(1);
    if (worksheet) {
      var cell = worksheet.Cells(1, 1);
      if (cell) {
        cell.value = "Hello world";
      }
    }
  },
  SheetActivate: function (worksheet) {
    console.log('SheetActivate event triggerred');
  },
  SheetDeactivate: function (worksheet) {
    console.log('SheetDeactivate event triggerred');
  },
  SheetSelectionChange: function (worksheet, range) {
    //var worksheet1 = range; // it's a bug now: args in reversal order
    //var range = worksheet;
    console.log('SheetSelectionChange event triggerred', range.Address);
  },
  SheetChange: function (worksheet, range) {
    //var worksheet1 = range; // it's a bug now: args in reversal order
    //var range = worksheet;
    console.log('SheetChange event triggerred', range.Address, range.Value);
  },
});

application.Visible = true;
console.log('Excel application has been open.');
console.log('--> Please create a new workbook to trigger event "NewWorkbook"');

setTimeout(function () {
  clearInterval(interval);
	application.Quit();
	winax.release(application);
}, 1000 * 60);

