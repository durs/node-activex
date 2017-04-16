# NAME

	Preview version Node.JS addon, that implements COM IDispatch object wrapper, analog ActiveXObject

# USAGE

	require('activex');
	var con = new ActiveXObject("ADODB.Connection");

	// another usage:
	// var ActiveX = require('activex');
	// var con = new ActiveX.Object("ADODB.Connection");
	
	var path = require('path'); 
	var dbf_path = path.join(__dirname, '../tmp');
	var dbf_file = "persons.dbf";
	var dbf_constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbf_path + ";Extended Properties=\"DBASE IV;\"";

	con.Open(dbf_constr, "", "");
	con.Execute("Create Table " + dbf_file + " (Name char(50), City char(50), Phone char(20), Zip decimal(5))");
	con.Execute("Insert into " + dbf_file + " Values('John', 'London','123-45-67','14589')");
	con.Execute("Insert into " + dbf_file + " Values('Andrew', 'Paris','333-44-55','38215')");
	con.Execute("Insert into " + dbf_file + " Values('Romeo', 'Rom','222-33-44','54323')");

	var rs = con.Execute("Select * from " + dbf_file); 
	var fields = rs.Fields;
	while (!rs.EOF) { 
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

