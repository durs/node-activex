var winax = require('../activex');

var path = require('path'); 
const assert = require('assert');

var js_arr = ['1', 2, 3];

describe("Variants", function() {

    it("Short Array", function() {
        var arr = new winax.Variant(js_arr, 'short');
        assert.equal(arr.length, js_arr.length);
        assert.strictEqual(arr[0], 1);
    });

    it("String Array", function() {
        var arr = new winax.Variant(js_arr, 'string');
        assert.equal(arr.length, js_arr.length);
        assert.strictEqual(arr[1], '2');
    });

    it("Variant Array", function() {
        var arr = new winax.Variant(js_arr, 'variant');
        assert.equal(arr.length, js_arr.length);
        assert.strictEqual(arr[0], js_arr[0]);
        assert.strictEqual(arr[1], js_arr[1]);
    });

    it("References", function() {
        var v = new winax.Variant();
        var ref = new winax.Variant(v, 'byref');

        v.assign(1, 'string');
        assert.strictEqual(v.valueOf(), '1');
        assert.strictEqual(ref.valueOf(), '1');

        v.cast('int');
        assert.strictEqual(v.valueOf(), 1);
        assert.strictEqual(ref.valueOf(), 1);

        v.clear();
        assert.strictEqual(v.valueOf(), undefined);
        assert.strictEqual(ref.valueOf(), undefined);
    });

});

