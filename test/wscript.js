const assert = require('assert');

function r(name) {
    const {execSync} = require('child_process')
    execSync(__dirname+'/wscript/nodewscript_test_cli.cmd '+__dirname+'/wscript/'+name+'.js', {timeout: 10000})
}

describe("Execute WScript Tests", function() {

    it('WScript Arguments', function (done) {
        r('arguments');
        done()
    })

    it('WScript Enumerator', function (done) {
        r('enumerator');
        done()
    })


    it('WScript Mixed', function (done) {
        r('mixed');
        done()
    })

    it('WScript WMI', function (done) {
        r('wmi');
        done()
    })


});

