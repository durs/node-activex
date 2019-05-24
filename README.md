# Name

	Windows C++ Node.JS addon, that implements COM IDispatch object wrapper, analog ActiveXObject on cscript.exe

# Features

 * Using *ITypeInfo* for conflict resolution between property and method 
 (for example !rs.EOF not working without type information, becose need define EOF as an object) 

 * Using optional parameters on constructor call    
``` js 
var con = new ActiveXObject("ADODB.Connection", {
	activate: false, // Allow activate existance object instance, false by default
	async: true, // Allow asynchronius calls, true by default (for future usage)
	type: true	// Allow using type information, true by default
});
```

 * Create COM object from JS object and may be send as argument (for example send to Excel procedure)
``` js 
var com_obj = new ActiveXObject({
	text: test_value,
	obj: { params: test_value },
	arr: [ test_value, test_value, test_value ],
	func: function(v) { return v*2; }
});
```

 * Additional COM variant types support:
	- **int** - VT_INT
	- **uint** - VT_UINT
	- **int8**, **char** - VT_I1
	- **uint8**, **uchar**, **byte** - VT_UI1
	- **int16**, **short** - VT_I2
	- **uint16**, **ushort** - VT_UI2
	- **int32** - VT_I4
	- **uint32** - VT_UI4
	- **int64**, **long** - VT_I8
	- **uint64**, **ulong** - VT_UI8
	- **currency** - VT_CY
	- **float** - VT_R4
	- **double** - VT_R8
	- **string** - VT_BSTR
	- **date** - VT_DATE
	- **decimal** - VT_DECIMAL
	- **variant** - VT_VARIANT
	- **null** - VT_NULL
	- **empty** - VT_EMPTY
	- **byref** - VT_BYREF or use prefix **'p'** to indicate reference to the current type

``` js 
var winax = require('winax');
var Variant = winax.Variant;

// create variant instance 
var v_short = new Variant(17, 'short');
var v_short_byref = new Variant(17, 'pshort');
var v_int_byref = new Variant(17, 'byref');
var v_byref = new Variant(v_short, 'byref');

// create variant arrays
var v_array_of_variant = new Variant([1,'2',3]);
var v_array_of_short = new Variant([1,'2',3], 'short');
var v_array_of_string = new Variant([1,'2',3], 'string');	

// change variant content
var v_test = new Variant();
v_test.assign(17);
v_test.cast('string');
v_test.clear();

// also may be used cast function
var v_short_from_cast = winax.cast(17, 'short');
```

 * Additional dignostic propeties:
	- **__id** - dispatch identity, for exmplae: ADODB.Connection.@Execute.Fields
	- **__value** - value of dispatch object, equiles valueOf()
	- **__type** - list member names with their properties

# Usage example

Install package throw NPM (see below **Building** for details)
```
npm install winax
npm install winax --msvs_version=2015
npm install winax --msvs_version=2017
```

Create ADO Connection throw global function
``` js
require('winax');
var con = new ActiveXObject('ADODB.Connection');
```
Or using Object prototype
``` js
var winax = require('winax');
var con = new winax.Object('ADODB.Connection');
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
Release COM objects (but other temporary objects may be keep references too)
``` js
winax.release(con, rs, fields)
```

# Tutorial and Examples

- [examples/ado.js](https://github.com/durs/node-activex/blob/master/examples/ado.js)

# Building

This project uses Visual C++ 2013 (or later versions then support C++11 standard).
Bulding also requires node-gyp and python 2.6 (or later) to be installed. 
You can do this with npm:
```
npm install --global --production windows-build-tools
```
To obtain and build use console commands:
```
git clone git://github.com/durs/node-axtivex.git
cd node-activex
node-gyp configure
node-gyp build
```

For Electron users, need rebuild with a different V8 version:
```
npm rebuild winax --runtime=electron --target=2.0.2 --disturl=https://atom.io/download/atom-shell --build-from-source
```
Change --target value to your electron version.
See also Electron Documentation: [Using Native Node Modules](https://electron.atom.io/docs/tutorial/using-native-node-modules/).

# Tests

[mocha](https://github.com/visionmedia/mocha) is required to run unit tests.
```
npm install -g mocha
mocha test
```

# Contributors

* [durs](https://github.com/durs)

