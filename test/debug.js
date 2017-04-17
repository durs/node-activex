
var ActiveX = require('../build/Debug/node_activex');

var path = require('path'); 
var tmp_path = path.join(__dirname, '../tmp/');

//var fso = new ActiveX.Object("Scripting.FileSystemObject");
//if (!fso.FolderExists(tmp_path)) fso.CreateFolder(tmp_path);
//if (fso.FileExists(tmp_path + dbf_file)) fso.DeleteFile(tmp_path + dbf_file);

var dbf_file = "persons.dbf";
var dbf_constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tmp_path + ";Extended Properties=\"DBASE IV;\"";
var con = new ActiveX.Object("ADODB.Connection");
con.Open(dbf_constr, "", "");
var rs = con.Execute("Select * from " + dbf_file); 
var fields = rs.Fields;
var cnt = fields.Count;
var name = fields("Name").Value;