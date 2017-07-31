//-------------------------------------------------------------------------------------------------------
// Project: node-activex
// Author: Yuri Dursin
// Description: Example of using ActiveX addon with ADO
//-------------------------------------------------------------------------------------------------------

//require('winax');
require('../activex');

var path = require('path'); 
var data_path = path.join(__dirname, '../data/');
var filename = "persons.dbf";
var constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + data_path + ";Extended Properties=\"DBASE IV;\"";

console.log("==> Preapre directory and delete DBF file on exists");
var fso = new ActiveXObject("Scripting.FileSystemObject");
if (!fso.FolderExists(data_path)) fso.CreateFolder(data_path);
if (fso.FileExists(data_path + filename)) fso.DeleteFile(data_path + filename);

console.log("==> Open connection");
var con = new ActiveXObject("ADODB.Connection");
console.log("ADO version: " + con.Version);
con.Open(constr, "", "");

console.log("==> Create new DBF file")
con.Execute("create Table " + filename + " (Name char(50), City char(50), Phone char(20), Zip decimal(5))");

console.log("==> Insert records to DBF")
con.Execute("insert into " + filename + " values('John', 'London','123-45-67','14589')");
con.Execute("insert into " + filename + " values('Andrew', 'Paris','333-44-55','38215')");
con.Execute("insert into " + filename + " values('Romeo', 'Rom','222-33-44','54323')");

console.log("==> Select records from DBF")
var rs = con.Execute("Select * from " + filename); 
var fields = rs.Fields;
console.log("Result field count: " + fields.Count);
console.log("Result record count: " + rs.RecordCount);

rs.MoveFirst();
while (!rs.EOF) {
    // Access as property by string key
    var name = fields["Name"].Value;

    // Access as method with string argument
    var town = fields("City").value;

    // Access as indexed array
    var phone = fields[2].value;
    var zip = fields[3].value;    

    console.log("> Person: "+name+" from " + town + " phone: " + phone + " zip: " + zip);
    rs.MoveNext();
}

con.Close();
