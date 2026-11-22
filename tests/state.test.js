const m = require(process.argv[2]);
const { state } = m;
const assert = require("assert");
let pass = 0, fail = 0;
const t = (name, fn) => { try { fn(); pass++; console.log("PASS", name); } catch (e) { fail++; console.log("FAIL", name, "\n   ", e.message.split("\n")[0]); } };
const seed = () => {
  state.lists = [{ id: "list_1", name: "Main" }, { id: "list_2", name: "Work" }];
  state.panes.left = { listId: "list_1", search: "", selectedSids: new Set() };
  state.panes.right = { listId: "list_2", search: "x", selectedSids: new Set() };
  state.columns = ["Name", "Phone", "Email"];
  state.colMap = { Name: "name", Phone: "phone", Email: "email" };
  state.visibleColumns = { Name: true, Phone: true, Email: true };
  state.match = { phone: true, name: true, email: false, org: false };
  state.loadedName = "a.csv";
  state.contacts = [m.hydrateContact({ values: { Name: "Ann Lee", Phone: "555-111-2222", Email: "a@x.com" }, listId: "list_1" }, 0)];
  m.computeAllDuplicates();
};
const dump = () => JSON.stringify({ l: state.lists, p: state.panes.left.listId, r: state.panes.right.listId, c: state.contacts.map(c => [c._sid, c.listId, c.values]), m: state.match, cols: state.columns });

t("S1 payload round-trips through applyLoadedData", () => {
  seed(); const before = dump();
  const data = JSON.parse(JSON.stringify(m.buildPayload()));
  state.contacts = []; state.lists = []; m.applyLoadedData(data);
  assert.strictEqual(dump(), before);
});
t("S2 applyLoadedData is atomic on failure", () => {
  seed(); const before = dump();
  const bad = { v: 8, columns: ["Name"], colMap: {}, lists: [{ id: "x", name: "X" }], panes: {}, contacts: [{ get values() { throw new Error("boom"); } }] };
  let threw = false;
  try { m.applyLoadedData(bad); } catch (e) { threw = true; }
  assert.strictEqual(dump(), before, "state changed despite failure (threw=" + threw + ")");
});
t("S3 undo restores previous contacts", () => {
  seed(); m.history.undo.length = 0; m.history.redo.length = 0;
  const before = JSON.stringify(state.contacts.map(c => c.values));
  m.pushHistory();
  state.contacts[0].values.Name = "Changed";
  m.undo();
  assert.strictEqual(JSON.stringify(state.contacts.map(c => c.values)), before);
});
t("S4 redo reapplies", () => {
  seed(); m.history.undo.length = 0; m.history.redo.length = 0;
  m.pushHistory(); state.contacts[0].values.Name = "Changed";
  m.undo(); m.redo();
  assert.strictEqual(state.contacts[0].values.Name, "Changed");
});
console.log(`\n${pass} passed, ${fail} failed`);
