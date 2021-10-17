const winax = require('../activex');

const path = require('path');
const assert = require('assert');
const x64 = process.arch.indexOf('64') >= 0;

const js_arr = ['1', 2, 3];

describe("Variants", function() {

    it("Short Array", function() {
        const arr = new winax.Variant(js_arr, 'short');
        assert.equal(arr.length, js_arr.length);
        assert.strictEqual(arr[0], 1);
    });

    it("String Array", function() {
        const arr = new winax.Variant(js_arr, 'string');
        assert.equal(arr.length, js_arr.length);
        assert.strictEqual(arr[1], '2');
    });

    it("Variant Array", function() {
        const arr = new winax.Variant(js_arr, 'variant');
        assert.equal(arr.length, js_arr.length);
        assert.strictEqual(arr[0], js_arr[0]);
        assert.strictEqual(arr[1], js_arr[1]);
    });

    it("References", function() {
        const v = new winax.Variant();
        const ref = new winax.Variant(v, 'byref');

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

    it("Dispatch pointer as Uint8Array", function() {

        // create simple dispatch object
        const obj = new winax.Object({
            test: function(v) { return v * 2; }
        });

        // convert dispatch pointer to bytes array
        const varr = new winax.Variant(obj, 'uint8[]');

        // convert to javascript Uint8Array and check length
        const arr = new Uint8Array(varr.valueOf());
        assert.strictEqual(arr.length, x64 ? 8 : 4);

        // create dispatch object from array pointer and test
        const obj2 = new winax.Object(arr);
        assert.strictEqual(obj2.test(2), 4);
    });

    if (typeof global.gc === 'function') {
        global.gc();
        const mem_usage = process.memoryUsage().heapUsed / 1024;
        it("Check memory", function() {
            global.gc();
            const mem = process.memoryUsage().heapUsed / 1024;
            if (mem > mem_usage) throw new Error(`used memory increased from ${mem_usage.toFixed(2)}Kb to ${mem.toFixed(2)}Kb`);
        });
    }
});
