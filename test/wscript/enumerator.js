//Reference: <https://groups.google.com/group/comp.lang.javascript/browse_thread/thread/684ad16518c837a2/67d00aa5dbe854c2?show_docid=67d00aa5dbe854c2>
var listFileTypes = (function(){
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var shell = new ActiveXObject("WScript.Shell");

    function isKnown(file){
        var fName = file.Name;
        //the rules of capitalization are strange in windows...
        //there are rare cases where this can fail
        //for example: .HKEY_CLASSES_ROOT\.HeartsSave-ms
        var ext = fName.slice(fName.lastIndexOf(".")).toLowerCase();

        try{
            shell.RegRead("HKCR\\"+ext+"\\");
            return "Yes"
        } catch(e){
            return "No"
        }
    }

    return function(folder){
        var files = new Enumerator(fso.GetFolder(folder).Files);
        for(;!files.atEnd();files.moveNext()){
            var file = files.item();
            WScript.Echo(file.Name + "\t" + file.Type + "\t" + isKnown(file));
        }
    }
})()


listFileTypes("C:\\temp")