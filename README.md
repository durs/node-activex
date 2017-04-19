# Name

	Windows C++ Node.JS addon, that implements COM IDispatch object wrapper, analog ActiveXObject on cscript.exe

# Features

 * Using *ITypeInfo* for conflict resolution between property and method 
 (for example !rs.EOF not working without type information, becose need define EOF as an object) 

 * Using optional parameters on constructor call    
``` js 
	var con = new ActiveX.Object("ADODB.Connection", {
		async: true, // Allow asynchronius calls, true by default (for future usage)
		type: true	// Allow using type information, true by default
	});
```

 * Additional dignostic propeties:
	- *__id* - dispatch identity, for exmplae: ADODB.Connection.@Execute.Fields
	- *__value* - value of dispatch object, equiles valueOf()
	- *__type* - full list type members names with their properties

# Perspectives

 * Convert a JavaScript object to IDispatch
 * Asynchronius calls
 * Dispose method

# Usage example

Create ADO Connection throw global function
``` js
	require('activex');
	var con = new ActiveXObject('ADODB.Connection');
```
Or using native Object prototype
``` js
	var ActiveX = require('activex');
	var con = new ActiveX.Object('ADODB.Connection');
```
Open connection and create simple table
``` js
	con.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\tmp;Extended Properties="DBASE IV;"', '', '');
	con.Execute("Create Table persons.dbf (Name char(50), City char(50), Phone char(20), Zip decimal(5))");
	con.Execute("Insert into persons.dbf Values('John', 'London','123-45-67','14589')");
	con.Execute("Insert into persons.dbf Values('Andrew', 'Paris','333-44-55','38215')");
	con.Execute("Insert into persons.dbf Values('Romeo', 'Rom','222-33-44','54323')");
```
Select query and return RecordSet object
``` js
	var rs = con.Execute("Select * from persons.dbf"); 
	var reccnt = rs.RecordCount;
```
Inspect RecordSet fields
``` js
	var fields = rs.Fields;
	var fldcnt = fields.Count;
```
Process records
``` js
	rs.MoveFirst();
	while (!rs.EOF) {
		var name = fields("Name").value;
		var town = fields("City").value;
		var phone = fields("Phone").value;
		var zip = fields("Zip").value;   
		console.log("> Person: " + name + " from " + town + " phone: " + phone + " zip: " + zip);    
		rs.MoveNext();
	}
```
Or using fields by index
``` js
	rs.MoveFirst();
	while (rs.EOF != false) {
		var name = fields[0].value;
		var town = fields[1].value;
		var phone = fields[2].value;
		var zip = fields[3].value;   
		console.log("> Person: " + name + " from " + town + " phone: " + phone + " zip: " + zip);    
		rs.MoveNext();
	}
```

# Building

This project uses Visual C++ 2013 (or later versions then support C++11 standard) and Python 2.6 (or later).
Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build use console commands:

    git clone git://github.com/durs/node-axtivex.git
    cd node-activex
    node-gyp configure
    node-gyp build

# CONTRIBUTORS

* [durs](https://github.com/durs)

