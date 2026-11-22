const m = require(process.argv[2]); const { state } = m; const assert = require("assert");
state.lists=[{id:"list_1",name:"Main"}]; state.columns=["Name","Surname","Phone 1 - Value","Phone 2 - Value","Email","Company"];
state.colMap={"Name":"name","Surname":"surname","Phone 1 - Value":"phone","Phone 2 - Value":"phone","Email":"email","Company":"org"};
const mk=(v,i)=>m.hydrateContact({values:v,listId:"list_1"},i);
state.contacts=[
  mk({Name:"Ann",Surname:"Lee","Phone 1 - Value":"+91 98223 22307","Phone 2 - Value":"","Email":"a@x.com",Company:""},0),
  mk({Name:"Ann",Surname:"Lee","Phone 1 - Value":"98223 22307","Phone 2 - Value":"+91 92712 12052","Email":"A@x.com",Company:"Acme"},1),
  mk({Name:"Bob",Surname:"Ray","Phone 1 - Value":"","Phone 2 - Value":"","Email":"",Company:"Zeta"},2)];
state.contacts.forEach((c,i)=>c.id=i);
state.contacts.forEach(c=>c._sid = c._sid);
const sids=[state.contacts[0]._sid,state.contacts[1]._sid];
assert.strictEqual(m.mergeRecords(sids), true);
assert.strictEqual(state.contacts.length, 2, "one contact removed");
const p = state.contacts[0];
console.log("phones:", JSON.stringify(p.values["Phone 1 - Value"]), "| phone2:", JSON.stringify(p.values["Phone 2 - Value"]), "| email:", JSON.stringify(p.values["Email"]), "| company:", JSON.stringify(p.values["Company"]));
assert.strictEqual(p.values["Company"], "Acme");
assert.strictEqual(p.values["Phone 2 - Value"], "");
console.log("PASS merge");
