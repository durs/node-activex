# Name

	Windows C++ Node.JS addon, that implements COM IDispatch object wrapper, analog ActiveXObject on cscript.exe

# Features

 * Using *ITypeInfo* for conflict resolution between property and method 
 (for example !rs.EOF not working without type information, becose need define EOF as an object) 

 * Using optional parameters on constructor call    
``` js 
var con = new ActiveXObject("Object.Name", {
	activate: false, // Allow activate existance object instance (through CoGetObject), false by default
	getobject: false, // Allow using name of the file in the ROT (through GetAccessibleObject), false by default
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
	- **__type** - array all member items with their properties
	- **__methods** - list member mathods by names (ITypeInfo::GetFuncDesc)
	- **__vars** - list member variables by names (ITypeInfo::GetVarDesc) 

 * Full WScript Emulation Support through nodewscript
	- **ActiveXObject** type
	- **WScript.CreateObject**
	- **WScript.ConnectObject**
	- **WScript.DisconnectObject**
	- **WScript.Sleep**
	- **WScript.Arguments**
	- **WScript.Version**
	- **GetObject**
	- **Enumerator**

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
Working with Excel ranges using two dimension arrays (from 1.18.0 version)
* The second dimension is only deduced from the first array.
* If data is missing at the time of SafeArrayPutElement, VT_EMPTY is used.
* Best way to explicitly pass VT_EMPTY is to use null
* If the SAFEARRAY dims are smaller than those of the range, the missing cells are emptied.
* Exception to 4! Excel may process the SAFEARRAY as it does when processing a copy/paste operation
``` js
var excel = new winax.Object("Excel.Application", { activate: true });
var wbk = excel.Workbooks.Add(template_filename);
var wsh = wbk.Worksheets.Item(1);
wsh.Range("C3:E4").Value = [ ["C3", "D3", "E3" ], ["C4", "D4", "E4" ] ];
wsh.Range("C3:E4").Value = [ ["C3", "D3", "E3" ], "C4" ]; // will let D4 and E4 empty
wsh.Range("C3:E4").Value = [ [null, "D3", "E3" ], "C4" ]; // will let C3, D4 and E4 empty
wsh.Range("C3:F4").Value = [ [100], [200] ]; // will duplicate the two rows in colums C, D, E, and F
wsh.Range("C3:F4").Value = [ [100, 200, 300, 400] ]; // will duplicate the for cols in rows 3 and 4
wsh.Range("C3:F4").Value = [ [100, 200] ]; // Will correctly duplicate the first two cols, but col E and F with contains "#N/A"
const data = wsh.Range("C3:E4").Value.valueOf();
console.log("Cell E4 value is", data[1][2]);
```

# Tutorial and Examples

- [examples/ado.js](https://github.com/durs/node-activex/blob/master/examples/ado.js)

# Usage as a library

The repo allows to re-use some of its code as a library for your own native node addon.

To include it, install it as a normal node-dependency and then add it
to the `"dependencies"` section of your `binding.gyp` file like this:
```
"dependencies": [
  "<!(node -p \"require.resolve('winax/lib_binding.gyp')\"):lib_node_activex"
]
```

In your code, you can then include it using:
```cpp
#include <node_activex.h>
```

This makes functions like `Variant2Value` or `Value2Variant` that translate between COM VARIANT and node types
available in your code.
Note that providing this library functionality is not the core target of this repo however,
so importing it eg. currently declares
all methods in the global namespace and opens the namespaces `v8` and `node`.

Check the source code [`src/disp.h`](src/disp.h) and [`src/utils.h`](src/utils.h) for details.

# Building

This project uses Visual C++ 2013 (or later versions then support C++11 standard).
Bulding also requires node-gyp and python 2.6 (or later) to be installed. 
Supported NodeJS Versions (x86 or x64): 10, 11, 12, 13, 14, 15 
You can do this with npm:
```
npm install --global --production windows-build-tools
```
To obtain and build use console commands:
```
git clone git://github.com/durs/node-axtivex.git
cd node-activex
npm install
```
or debug version
```
npm install --debug
```
or using node-gyp directly
```
node-gyp configure
node-gyp build
```

For Electron users, need rebuild with a different V8 version:
```
npm rebuild winax --runtime=electron --target=12.22.0 --dist-url=https://electronjs.org/headers --build-from-source
```
Change --target value to your electron version.
See also Electron Documentation: [Using Native Node Modules](https://electron.atom.io/docs/tutorial/using-native-node-modules/).

# WScript

[WScript](https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/wscript) is designed as a runtime, so WSH scripts should be executed from command line. WScript emulation mode allows one to use standard **WScript** features and mix them with nodejs and ES6 parts.

There is one significant difference between WScript and NodeJS in general. NodeJS is designed to be [non-blocking](https://nodejs.org/en/docs/guides/blocking-vs-non-blocking/) and execution is done when last statement is done AND all pending promises and callbacks are resolved. WScript runtime is linear. It is done when last statement is executed.

So NodeJS is more convenient for writing things like async web services. Also with NodeJS you get access to invaluable collection of **npm** modules.

WScript has impressive collection of Windows-specific APIs available through ActiveX (such as `WScript.Shell`, `FileSystemObject` etc). Also it supports application events to implement two way communication with external apps and processes. 

This module allows taking the best from the two worlds - you may now use **npm** modules in your WScript scripts. Or, if you have bunch of legacy JScript sources, you may now execute them using this **NodeJS**-based runtime.

If you install this package with -g key, you will have `nodewscript` command. 

Usage:
```
nodewscript [options] <Filename.js>
```

Its direct analog in the windows scripting host is `cscript.exe` (it it outputs to the console, and `WScript.Echo` writes a line instead of showing a popup window).

Where `Filename.js` is a file designed for windows scripting host.

## WScript Limitations

### Implicit Properties

This is a key drawback of v8 engine compared to MS. Consider an example:

```javascript
    var a = WScript.CreateObject(...)
    if( a.Prop )
    {
        // If 'Prop' is a dynamic property (i.e. it is not defined in TypeInfo and
        // not marked explicitly as a property, then execution never gets here. Even if a.Prop is null or false.
    }
```

### Function Setters

```javascript
var WshShell = new ActiveXObject("WScript.Shell");
var processEnv = WshShell.Environment("PROCESS");
processEnv("ENV_VAR") = "CUSTOM_VALUE";  // This syntax is valid for JScript, but will throw syntax error on v8
```
### Events & Sleep

Default WScript engine checks for events every time you do some sync operation or Sleep. In this implementation we check for events when `WScript.Sleep` is executed. 

### 32 vs 64 Bit

It is using the same bitness as installed version of the  nodejs. So if your JScript files are designed for 32 bit, make sure to have nodejs x86 version installed before installing the **winax**.


# Tests

[mocha](https://github.com/visionmedia/mocha) is required to run unit tests.
```
npm install -g mocha
mocha --expose-gc test
```

# Contributors

* [durs](https://github.com/durs)
* [somanuell](https://github.com/somanuell)
* [Daniel-Userlane](https://github.com/Daniel-Userlane)
* [alexeygrinevich](https://github.com/alexeygrinevich)

