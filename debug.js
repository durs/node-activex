//-------------------------------------------------------------------------------------------------------
// Project: node-activex
// Author: Yuri Dursin
// Description: Script for addon debugging purposes
//-------------------------------------------------------------------------------------------------------

var path = require('path'); 
var tmp_path = path.join(__dirname, './tmp/');

function debug() {
    
    var ActiveX = require('./build/Debug/node_activex');
    
    var js_obj = {
        text: 'dispatch object from javascript',
        lines: ['dispatch', 'object', 'from', 'javascript'],
        conf: { params: 10 },
        getText: function(i) {
            if (arguments.length == 0) return text;
            return this.lines[i];
        }
    };
    var com_obj = new ActiveX.Object(js_obj);    
    com_obj.text = 'no text';

    //!!! Don`t work set by index
    com_obj.lines[0] = 'unknown';
    //var com_obj_lines = com_obj.lines;
    //var com_obj_lines_0 = com_obj.lines[0];

    //var my_arr_item = com_obj.lines[0].__value;
    //var my_obj_prop = com_obj.conf.params.__value;
    //var my_func_call = com_obj.getText(3);
    
    var excel = new ActiveX.Object("Excel.Application", { activate: true });
    if (!excel.Visible) excel.Visible = true;
    var wbk = excel.Workbooks.Add(tmp_path + 'test.xltm');

    var test1 = wbk.Test(com_obj, 'text', 0, 0);
    var test2 = wbk.Test(com_obj, 'getText()', 0, 0);
    var test3 = wbk.Test(com_obj, 'lines[]', 0, 0);
    var test4 = wbk.Test(com_obj, 'lines[]=', 0, 'test');

    //var fso = new ActiveX.Object("Scripting.FileSystemObject");
    //if (!fso.FolderExists(tmp_path)) fso.CreateFolder(tmp_path);
    //if (fso.FileExists(tmp_path + dbf_file)) fso.DeleteFile(tmp_path + dbf_file);

    var dbf_file = "persons.dbf";
    var dbf_constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tmp_path + ";Extended Properties=\"DBASE IV;\"";
    var con = new ActiveX.Object("ADODB.Connection", {
        async: true,
        type: true
    });
    con.Open(dbf_constr, "", "");
    var rs = con.Execute("Select * from " + dbf_file); 
    var fields = rs.Fields;
    var cnt = fields.Count;

    // Test identity
    //var identity = fields.__id;

    // Test type info
    //var type = fields.__type;

    // Query by index
    var str1 = fields[0].Value;    

    // Not worked!!!
    //var str2 = rs("Name");

	rs.MoveFirst();
	while (rs.EOF != true) { 
		var name = fields("Name").value;
		var town = fields("City").value;
		var phone = fields("Phone").value;
		var zip = fields("Zip").value;   
		console.log("> Person: "+name+" from " + town + " phone: " + phone + " zip: " + zip);    
		rs.MoveNext();
	}

}

try {
    debug();
}
catch(e) {
    console.error(e);
}