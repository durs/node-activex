// This test should have both node and WScript mixed

// First, we get count of files using FSO

var fso = new ActiveXObject("Scripting.FileSystemObject");

var filesCount = fso.GetFolder(".").Files.Count;

WScript.Echo("FSO Path: "+fso.GetFolder(".").Path+" Files count: "+filesCount);

// This if makes sure you still may run this file using cscript
if(WScript.Version=='NODE.WIN32')
{
    // And then use node fs
    console.log("path.resolve(.): "+require('path').resolve('.'));

    var fs = require('fs');
    var files = fs.readdirSync(".");

    var nodeFilesCount = 0;
    for(var f in files) 
    {
        if( !fs.lstatSync(files[f]).isDirectory() )
        {
            nodeFilesCount++;
        }
    }

    if(nodeFilesCount!=filesCount)
    {
        console.log("FSO files: ", filesCount, "fs files: ", nodeFilesCount);
        WScript.Quit(-5);
    } else {
        console.log("Same number of files: "+nodeFilesCount);
    }
}
