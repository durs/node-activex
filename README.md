# NAME

Preview version Node.JS addon, that implements COM IDispatch object wrapper, analog ActiveXObject

# USAGE

	require('activex');
	var con = new ActiveXObject("ADODB.Connection");

	// another usage:
	// var ActiveX = require('activex');
	// var con = new ActiveX.Object("ADODB.Connection");
	
	var dbf_path = "c:\\temp\\";
	var dbf_file = "persons.dbf";
	var dbf_constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbf_path + ";Extended Properties=\"DBASE IV;\"";

	con.Open(dbf_constr, "", "");
	con.Execute("Create Table " + dbf_file + " (Name char(50), City char(50), Phone char(20), Zip decimal(5))");
	con.Execute("Insert into " + dbf_file + " Values('John', 'London','123-45-67','14589')");
	con.Execute("Insert into " + dbf_file + " Values('Andrew', 'Paris','333-44-55','38215')");
	con.Execute("Insert into " + dbf_file + " Values('Romeo', 'Rom','222-33-44','54323')");

	var rs = con.Execute("Select * from " + dbf_file); 
	var fields = rs.Fields;
	while (!rs.EOF.valueOf() /*???*/)
	{ 
	    var name = fields[0].value;
		var town = fields[1].value;
		var phone = fields[2].value;
		var zip = fields[3].value;
		console.log("> Person: "+name+" from " + town + " phone: " + phone + " zip: " + zip);    
		rs.MoveNext();
	}

# Tutorial and Examples


# Other built in functions



# FEATURES


# API


# BUILDING

This project uses Visual C++ 2010 (or later) and Python 2.6 (or later).
Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build the bindings:

    git clone git://github.com/durs/node-axtivex.git
    cd node-activex
    node-gyp configure
    node-gyp build

# TESTS

[mocha](https://github.com/visionmedia/mocha) is required to run unit tests.

    npm install -g mocha
    nmake /a test


# CONTRIBUTORS

* [durs](https://github.com/durs)


# ACKNOWLEDGEMENTS

Inspired [Win32OLE](http://www.ruby-doc.org/stdlib/libdoc/win32ole/rdoc/)


# LICENSE

`node-win32ole` is [BSD licensed](https://github.com/idobatter/node-win32ole/raw/master/LICENSE).
