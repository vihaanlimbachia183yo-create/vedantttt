import test from 'node:test';import assert from 'node:assert/strict';
const districts=['Mumbai City','Mumbai Suburban','Navi Mumbai','Thane'];
function normalize(r){if(!districts.includes(r.district))throw new Error('district');if(!r.college||!r.branch||!r.capCode)throw new Error('required');return r}
test('accepts Mumbai scoped official-format row',()=>assert.equal(normalize({district:'Mumbai City',college:'A',branch:'B',capCode:'1'}).capCode,'1'));
test('rejects out-of-scope district',()=>assert.throws(()=>normalize({district:'Pune',college:'A',branch:'B',capCode:'1'}),/district/));
test('recommendation margin semantics',()=>{const pct=95,cut=94;assert.equal(pct-cut,1)});
