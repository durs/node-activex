# NAME

	Preview version Node.JS addon, that implements COM IDispatch object wrapper, analog ActiveXObject

# USAGE

	require('activex');
	var con = new ActiveXObject("ADODB.Connection");

	// another usage:
	// var ActiveX = require('activex');
	// var con = new ActiveX.Object("ADODB.Connection");
	
	var path = require('path'); 
	var dbf_path = path.join(__dirname, '../tmp/');
	var dbf_file = "persons.dbf";
	var dbf_constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" 
		+ dbf_path + ";Extended Properties=\"DBASE IV;\"";

	console.log("Preapre directory and delete DBF file on exists");
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	if (!fso.FolderExists(dbf_path)) fso.CreateFolder(dbf_path);
	if (fso.FileExists(dbf_path + dbf_file)) fso.DeleteFile(dbf_path + dbf_file);

	console.log("Open connection");
	var con = new ActiveXObject("ADODB.Connection");
	console.log("ADO version: " + con.Version);
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
	console.log("Resukt record count: " + rs.RecordCount);
	console.log("Resukt field count: " + fields.Count);

	rs.MoveFirst();
	while (!rs.EOF.valueOf()) { 
		var name = fields("Name").value;
		var town = fields("City").value;
		var phone = fields("Phone").value;
		var zip = fields("Zip").value;   
		console.log("> Person: "+name+" from " + town + " phone: " + phone + " zip: " + zip);    
		rs.MoveNext();
	}

# BUILDING

This project uses Visual C++ 2010 (or later) and Python 2.6 (or later).
Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build the bindings:

    git clone git://github.com/durs/node-axtivex.git
    cd node-activex
    node-gyp configure
    node-gyp build

# CONTRIBUTORS

* [durs](https://github.com/durs)

