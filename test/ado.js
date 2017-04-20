require('../activex');

var path = require('path'); 
const assert = require('assert');

var data_path = path.join(__dirname, '../data/');
var filename = "persons.dbf";
var constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + data_path + ";Extended Properties=\"DBASE IV;\"";
var fso, con, rs, fields, reccnt;

describe("Scripting.FileSystemObject", function() {

    it("create", function() {
        fso = new ActiveXObject("Scripting.FileSystemObject");
    });

    it("create data folder if not exists", function() {
        if (fso) {
            if (!fso.FolderExists(data_path)) 
                fso.CreateFolder(data_path);
        }
    });

    it("delete DBF file if exists", function() {
        if (fso) {
            if (fso.FileExists(data_path + filename)) 
                fso.DeleteFile(data_path + filename);
        }
    });

});

describe("ADODB.Connection", function() {

    it("create and open", function() {
        con = new ActiveXObject("ADODB.Connection");
        con.Open(constr, "", "");
    });

    it("create and fill table", function() {
        if (con) {
            con.Execute("create Table " + filename + " (Name char(50), City char(50), Phone char(20), Zip decimal(5))");
            con.Execute("insert into " + filename + " values('John', 'London','123-45-67','14589')");
            con.Execute("insert into " + filename + " values('Andrew', 'Paris','333-44-55','38215')");
            con.Execute("insert into " + filename + " values('Romeo', 'Rom','222-33-44','54323')");
            reccnt = 3;
        }
    });

    it("select records from table", function() {
        if (con) {
            var rs = con.Execute("Select * from " + filename); 
            var fields = rs.Fields;
        }
    });

    it("loop by records", function() {
        if (rs && fields) {
            var cnt = 0;
            rs.MoveFirst();
            while (!rs.EOF) { 
                cnt++;
                var name = fields("Name").Value;
                var town = fields("City").value;
                var phone = fields("Phone").value;
                var zip = fields("Zip").value;    
                rs.MoveNext();
            }           
            assert.equal(cnt, reccnt); 
        }
    });

});