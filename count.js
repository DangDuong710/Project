// // Get the active document
var doc = app.activeDocument;
var logFile = File('D:/thangvt21/pet_project/projects' + '/' + 'hoodie_01.csv');

logFile.open('a');
if (doc.hasOwnProperty('countItems') && doc.countItems.length > 0) {
    for (var i = 0; i < doc.countItems.length; i++) {
        var countGroup = doc.countItems[i];
        data = countGroup.position;
        logFile.writeln(data);
    }
}
logFile.close();
