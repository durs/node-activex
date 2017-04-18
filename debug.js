//-------------------------------------------------------------------------------------------------------
// Project: node-activex
// Author: Yuri Dursin
// Description: Script for addon debugging purposes
//-------------------------------------------------------------------------------------------------------

var path = require('path'); 
var tmp_path = path.join(__dirname, './tmp/');

function debug() {
    var ActiveX = require('./build/Debug/node_activex');

    //var fso = new ActiveX.Object("Scripting.FileSystemObject");
    //if (!fso.FolderExists(tmp_path)) fso.CreateFolder(tmp_path);
    //if (fso.FileExists(tmp_path + dbf_file)) fso.DeleteFile(tmp_path + dbf_file);

    var dbf_file = "persons.dbf";
    var dbf_constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tmp_path + ";Extended Properties=\"DBASE IV;\"";
    var con = new ActiveX.Object("ADODB.Connection", {
        async: true,
        type: false
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