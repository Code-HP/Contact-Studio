const path = process.argv[2];
const m = require(path);
const { state } = m;
const assert = require("assert");
let pass = 0, fail = 0;
const t = (name, fn) => { try { fn(); pass++; console.log("PASS", name); } catch (e) { fail++; console.log("FAIL", name, "\n   ", e.message); } };

state.lists = [{ id: "list_1", name: "Main" }];
state.columns = ["Name", "Surname", "Phone", "Email"];
state.colMap = { Name: "name", Surname: "surname", Phone: "phone", Email: "email" };
state.match = { phone: true, name: true, email: true };
const mk = (name, surname, phone, email) => {
  const c = m.hydrateContact({ values: { Name: name, Surname: surname, Phone: phone, Email: email }, listId: "list_1" }, 0);
  return c;
};
const dupes = (contacts) => {
  state.contacts = contacts; contacts.forEach((c, i) => (c.id = i));
  return m.computeDuplicatesForList("list_1");
};
const fmt = (v, label) => m.formatPhoneCellValue(v, label);

// ---- Formatter ----
t("F1 bare number gets row label", () => assert.strictEqual(fmt("5551234567", "Cell"), "Cell: 5551234567"));
t("F2 label words map", () => assert.strictEqual(fmt("Mobile 555-123-4567; Work: 555 987 6543", ""), "Cell: 555-123-4567, Work: 555 987 6543"));
t("F3 cross-format dup removed", () => assert.strictEqual(fmt("+1 555 123 4567, 555-123-4567", ""), "Main: +1 555 123 4567"));
t("F4 trailing separator tidied", () => assert.strictEqual(fmt("555-1234-", ""), "Main: 555-1234"));
t("F5 label after number", () => assert.strictEqual(fmt("555-1234 (Home)", ""), "Home: 555-1234"));
t("F6 sequence exhaustion", () => console.log("    info:", fmt("1, 2, 3, 4, 5, 6", "")));

// ---- Duplicate finder ----
t("D1 reversed name order matches", () => {
  const g = dupes([mk("John", "Smith", "555-111-2222", ""), mk("Smith", "John", "555-999-0000", "")]);
  assert.strictEqual(g.length, 1, "expected 1 group, got " + g.length);
});
t("D2 placeholder phones do not group strangers", () => {
  const g = dupes([mk("Ann", "Lee", "0000000000", ""), mk("Bob", "Ray", "000-000-0000", "")]);
  assert.strictEqual(g.length, 0, "expected 0 groups, got " + g.length);
});
t("D3 same email groups", () => {
  const g = dupes([mk("Ann", "Lee", "555-111-2222", "A@x.com"), mk("Anne", "Leigh", "555-333-4444", "a@x.COM")]);
  assert.strictEqual(g.length, 1, "expected 1 group, got " + g.length);
});
t("D4 reason label is lowercase key", () => {
  const g = dupes([mk("John", "Smith", "555-111-2222", ""), mk("John", "Smith", "555-999-0000", "")]);
  assert.deepStrictEqual(g[0].reasons, ["name"]);
});
t("D5 phone-only group reasons", () => {
  const g = dupes([mk("A", "B", "+1 555 111 2222", ""), mk("C", "D", "5551112222", "")]);
  assert.deepStrictEqual(g[0].reasons, ["phone"]);
});
console.log(`\n${pass} passed, ${fail} failed`);
