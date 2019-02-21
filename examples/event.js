var winax = require('..');
setInterval(function () {
	winax.peekAndDispatchMessages();
}, 1000);

var application = new ActiveXObject('Excel.Application');

var connectionPoints = winax.getConnectionPoints(application);
var connectionPoint = connectionPoints[0];

var first = true;
connectionPoint.advise(function(workbook) {
	if (!first) return; // actually 'NewWorkbook' event
	first = false;
	console.log('event triggerred');
	var worksheet = workbook.Worksheets(1);
	if (worksheet) {
		var cell = worksheet.Cells(1,1);
		if (cell) {
			cell.value = "Hello world";
		}
	}
	
});

application.Visible = true;
console.log('Excel application has been open.\nPlease create a new workbook to trigger event "NewWorkbook"');

setTimeout(function () {
	application.Quit();
	winax.release(application);
}, 1000 * 60);

