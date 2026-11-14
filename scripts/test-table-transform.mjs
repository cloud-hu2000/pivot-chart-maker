import assert from 'node:assert/strict';
import {transposeTable,mergeTables} from '../src/table-transform.js';
assert.deepEqual(transposeTable([['a','b','c'],['1','2']]),[['a','1'],['b','2'],['c','']]);
assert.deepEqual(mergeTables([[['id','name'],['1','Ada']],[['name','id','age'],['Sam','2','30']]]),[['id','name','age'],['1','Ada',''],['2','Sam','30']]);
assert.deepEqual(mergeTables([[['k'],['x'],['x']],[['k'],['x']]]),[['k'],['x'],['x'],['x']]);
assert.throws(()=>mergeTables([[['a','a'],['x','y']],[['a'],['z']]]));
assert.throws(()=>mergeTables([[['a','b'],['x']],[['a'],['z']]]));
assert.throws(()=>mergeTables([[[''],['x']],[['a'],['z']]]));
console.log('Transpose and merge checks passed');
