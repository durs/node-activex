//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Last Modification: 2011-11-22
// Description: Example of using ActiveX addon for NodeJS
//-------------------------------------------------------------------------------------------------------

require('./Release/activex'); 
//require('./Debug/activex');

var dbf_path = "c:\\temp\\";
var dbf_file = "persons.dbf";
var dbf_constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbf_path + ";Extended Properties=\"DBASE IV;\"";

console.log("Preapre directory and delete DBF file on exists");
var fso = new ActiveXObject("Scripting.FileSystemObject");
if (!fso.FolderExists(dbf_path)) fso.CreateFolder(dbf_path);
if (fso.FileExists(dbf_path + dbf_file)) fso.DeleteFile(dbf_path + dbf_file);

console.log("Open connection");
var con = new ActiveXObject("ADODB.Connection");
con.Open(dbf_constr, "", "");

console.log("Create new DBF file")
con.Execute("Create Table " + dbf_file + " (Name char(50), City char(50), Phone char(20), Zip decimal(5))");

console.log("Insert records to DBF")
con.Execute("Insert into " + dbf_file + " Values('John', 'London','123-45-67','14589')");
con.Execute("Insert into " + dbf_file + " Values('Andrew', 'Paris','333-44-55','38215')");
con.Execute("Insert into " + dbf_file + " Values('Romeo', 'Rom','222-33-44','54323')");

console.log("Select records from DBF")
var rs = con.Execute("Select * from " + dbf_file); 
var fields = rs.Fields;

//while (!rs.EOF)
while (!rs.EOF.valueOf()/*???*/)
{ 
    //name=  Persons.Fields("Name").value;
    //town=  Persons.Fields("City").value;
    //phone=Persons.Fields("Phone").value;
    //zip=    Persons.Fields("Zip").value;    

    var name = fields[0].value;
    var town = fields[1].value;
    var phone = fields[2].value;
    var zip = fields[3].value;

    console.log("> Person: "+name+" from " + town + " phone: " + phone + " zip: " + zip);
    
    rs.MoveNext();
}

con.Close();
