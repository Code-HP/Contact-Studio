window.__CONTACT_STUDIO_XLSX_READY__ = typeof XLSX !== "undefined";
(function () {
  "use strict";

  /* ========== 1. STATE ========== */
  const state = {
    lists: [{ id: "list_1", name: "Main" }],
    panes: {
      left: { listId: "list_1", search: "", selectedSids: new Set() },
      right: { listId: null, search: "", selectedSids: new Set() },
    },
    contacts: [],
    columns: [],
    colMap: {},
    visibleColumns: {},
    groupsByList: new Map(),
    groupColors: new Map(),
    match: { phone: true, name: true, email: true },
    onlyNoNumber: false,
    onlySelfDup: false,
    activeGroup: null, // { id, listId } when a duplicate group is shown alone
    loadedName: "",
    editingId: null,
    editingListId: null,
    detailsContactSid: null,
  };

  /* ========== 2. CONSTANTS ========== */
  const SESSION_PREFIX =
    "s" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  let sidCounter = 0;
  const makeSid = () => SESSION_PREFIX + "_" + ++sidCounter;
  const STORAGE_KEY = "cdf.state.v8";
  const STORAGE_VERSION = 8;
  const HISTORY_LIMIT = 60;
  const DEFAULT_COLUMNS = [
    "Name",
    "Surname",
    "Phone",
    "Email",
    "Company",
    "Notes",
  ];
  const MATCH_TYPES = ["phone", "name", "email"];
  const AUTO_LABEL_SEQUENCE = ["Main", "Work", "Home", "Other"];
  const REASON_LABEL = { phone: "Phone", name: "Name", email: "Email" };
  // Digit strings that are almost always placeholders, not real numbers.
  const PLACEHOLDER_PHONE_SEQ = ["1234567890", "0123456789", "9876543210"];
  const XLSX_EXTS = ["xlsx", "xls", "xlsm", "xlsb", "ods"];
  const IDB_NAME = "contact_studio";
  const IDB_STORE = "state";
  const IDB_KEY = "main";

  const PHONE_LABEL_MAP = {
    main: "MAIN",
    primary: "MAIN",
    mainline: "MAIN",
    preferred: "MAIN",
    pref: "MAIN",
    work: "WORK",
    office: "WORK",
    business: "WORK",
    direct: "WORK",
    home: "HOME",
    house: "HOME",
    personal: "HOME",
    private: "HOME",
    cell: "CELL",
    mobile: "CELL",
    mobilephone: "CELL",
    gsm: "CELL",
    fax: "FAX",
    faxnumber: "FAX",
    other: "OTHER",
    alt: "OTHER",
    alternative: "OTHER",
    whatsapp: "WHATSAPP",
    wa: "WHATSAPP",
    telegram: "TELEGRAM",
    pager: "PAGER",
    voice: "VOICE",
  };
  const VCF_TYPE_MAP = {
    MAIN: "VOICE,PREF",
    WORK: "WORK,VOICE",
    HOME: "HOME,VOICE",
    CELL: "CELL,VOICE",
    FAX: "FAX",
    PAGER: "PAGER",
    OTHER: "VOICE",
    ALT: "VOICE",
    WHATSAPP: "CELL,VOICE",
    TELEGRAM: "CELL,VOICE",
    VOICE: "VOICE",
  };
  const HEADER_HINTS =
    /^(id|#|no|index|name|full ?name|display ?name|contact ?name|formatted ?name|complete ?name|first ?name|given ?name|forename|middle ?name|last ?name|family ?name|surname|nick ?name|nickname|alias|prefix|suffix|salutation|honorific|e-?mail|e-?mail ?address|email|email ?address|mail|phone|phones|phone ?no|phone ?number|mobile|mobile ?no|mobile ?number|cell|cellular|tel|telephone|work ?phone|home ?phone|personal ?phone|primary ?phone|secondary ?phone|other ?phone|fax|landline|whatsapp|msisdn|company|company ?name|organi[sz]ation|organi[sz]ation ?name|org|employer|business|account ?name|work|workplace|firm|title|job ?title|position|designation|role|department|dept|profession|occupation|address|street|city|state|country|zip|postal|notes?|comment|remark|description|memo|website|url|group|tags?|categories|label|labels)$/i;

  /* ========== 3. DOM ========== */
  const $ = (id) => document.getElementById(id);
  const els = {
    fileInput: $("fileInput"),
    dropOverlay: $("dropOverlay"),
    resetMemoryBtn: $("resetMemoryBtn"),
    toggleDupesBtn: $("toggleDupesBtn"),
    saveNowBtn: $("saveNowBtn"),
    saveStatus: $("saveStatus"),
    dupesDrawer: $("dupesDrawer"),
    dupesClose: $("dupesClose"),
    dupesList: $("dupesList"),
    matchPhone: $("matchPhone"),
    matchName: $("matchName"),
    matchEmail: $("matchEmail"),
    onlyNoNumber: $("onlyNoNumber"),
    onlySelfDup: $("onlySelfDup"),
    contactModal: $("contactModal"),
    modalTitle: $("modalTitle"),
    modalBody: $("modalBody"),
    modalForm: $("modalForm"),
    modalClose: $("modalClose"),
    modalCancel: $("modalCancel"),
    detailsModal: $("detailsModal"),
    detailsModalTitle: $("detailsModalTitle"),
    detailsBody: $("detailsBody"),
    detailsModalClose: $("detailsModalClose"),
    detailsModalCloseBtn: $("detailsModalCloseBtn"),
    detailsModalEditBtn: $("detailsModalEditBtn"),
    detailsModalDeleteBtn: $("detailsModalDeleteBtn"),
    detailsModalDedupeBtn: $("detailsModalDedupeBtn"),
    detailsModalLabelBtn: $("detailsModalLabelBtn"),
  };
  const paneEls = (side) => ({
    wrap: document.querySelector('[data-pane="' + side + '"]'),
    count: document.querySelector('[data-pane-count="' + side + '"]'),
    search: document.querySelector('[data-pane-search="' + side + '"]'),
    addBtn: document.querySelector('[data-add-pane="' + side + '"]'),
    exportBtn: document.querySelector('[data-export-pane="' + side + '"]'),
    exportMenu: document.querySelector('[data-export-menu="' + side + '"]'),
    selBar: document.querySelector('[data-pane-sel="' + side + '"]'),
    selCount: document.querySelector('[data-pane-sel-count="' + side + '"]'),
    delSelBtn: document.querySelector('[data-del-sel="' + side + '"]'),
    clearSelBtn: document.querySelector('[data-clear-sel="' + side + '"]'),
    body: document.querySelector('[data-pane-body="' + side + '"]'),
    thead: document.querySelector('[data-pane-thead="' + side + '"]'),
    tbody: document.querySelector('[data-pane-tbody="' + side + '"]'),
    hint: document.querySelector('[data-pane-hint="' + side + '"]'),
  });

  /* ========== 4. UTILITIES ========== */
  const stripBOM = (t) => (t && t.charCodeAt(0) === 0xfeff ? t.slice(1) : t);
  const escapeHtml = (s) =>
    String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  const normalizeHeader = (s) =>
    String(s || "")
      .trim()
      .toLowerCase()
      .replace(/[_\-\.]+/g, " ")
      .replace(/\s+/g, " ");
  const stripColSuffix = (n) =>
    String(n || "")
      .replace(/\s*[#(]?\s*\d+\s*[)]?\s*$/g, "")
      .replace(/\s+/g, " ")
      .trim();
  const splitList = (v) =>
    String(v || "")
      .split(/\s*:::\s*|[;|\/\n,]+|\s{3,}/)
      .map((s) => s.trim())
      .filter(Boolean);
  // Indian mobile numbers are shown as "+91 92712 12052": country code, then the
  // 10-digit number in two groups of five. Accepts a bare 10-digit number, a
  // leading 0 (0 95-276-22377) or a +91 / 91 prefix. Other numbers are unchanged.
  function addIndiaCode(n) {
    const s = String(n || "").trim();
    const d = s.replace(/\D+/g, "");
    if (!d) return s;
    let ten = "";
    if (s.startsWith("+")) {
      if (d.length === 12 && d.startsWith("91")) ten = d.slice(2);
      else return s;
    } else if (d.length === 10) ten = d;
    else if (d.length === 11 && d[0] === "0") ten = d.slice(1);
    else if (d.length === 12 && d.startsWith("91")) ten = d.slice(2);
    if (ten.length === 10 && /^[6-9]/.test(ten))
      return "+91 " + ten.slice(0, 5) + " " + ten.slice(5);
    return s;
  }
  function cleanCellValue(v) {
    if (v == null) return "";
    if (typeof v === "number") {
      if (Number.isInteger(v)) return String(v);
      const asInt = Math.trunc(v);
      return asInt === v ? String(asInt) : String(v);
    }
    if (typeof v === "boolean") return v ? "true" : "false";
    return String(v)
      .trim()
      .replace(/^(-?\d+)\.0+$/, "$1");
  }
  const normPhone = (v) => {
    const d = String(v || "").replace(/\D+/g, "");
    if (d.length < 7) return "";
    return d.length > 10 ? d.slice(-10) : d;
  };
  const normEmail = (v) =>
    String(v || "")
      .trim()
      .toLowerCase()
      .replace(/^mailto:/, "");
  const normName = (v) =>
    String(v || "")
      .trim()
      .toLowerCase()
      .replace(/[^\p{L}\p{N}]+/gu, " ")
      .replace(/\s+/g, " ")
      .trim();
  const normNamePart = (v) => {
    const s = normName(v);
    return s && s.length >= 2 ? s : "";
  };
  // Phone key used for duplicate matching. Returns "" for placeholder numbers
  // (e.g. 0000000000, 1234567890) so unrelated people are never grouped.
  const normPhoneKey = (v) => {
    const k = normPhone(v);
    if (!k) return "";
    if (/^(\d)\1+$/.test(k) || PLACEHOLDER_PHONE_SEQ.indexOf(k) !== -1) return "";
    return k;
  };
  // Name key that ignores word order and punctuation: "Smith John" == "John Smith".
  const nameMatchKey = (v) => {
    const s = normNamePart(v);
    return s ? s.split(" ").sort().join(" ") : "";
  };
  function hasRepeatedKey(list, normFn) {
    if (!list || list.length < 2) return false;
    const seen = new Set();
    for (const item of list) {
      const k = normFn(item);
      if (!k) continue;
      if (seen.has(k)) return true;
      seen.add(k);
    }
    return false;
  }
  const listName = (listId) => {
    const l = state.lists.find((x) => x.id === listId);
    return l ? l.name : "list";
  };

  /* ========== 5. PERSISTENT STORAGE ========== */
  let storageWarned = false;
  let saveTimer = null;
  let statusTimer = null;

  const localStorageOk = (() => {
    try {
      const t = "__cdf__";
      localStorage.setItem(t, t);
      localStorage.removeItem(t);
      return true;
    } catch (e) {
      return false;
    }
  })();
  const indexedDbOk = typeof indexedDB !== "undefined" && indexedDB !== null;

  function idbOpen() {
    return new Promise((resolve, reject) => {
      if (!indexedDbOk) {
        reject(new Error("IndexedDB unavailable"));
        return;
      }
      const req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(IDB_STORE))
          db.createObjectStore(IDB_STORE);
      };
      req.onsuccess = () => {
        const db = req.result;
        // If another tab asks to delete the database (Reset), let go of it at once.
        db.onversionchange = () => db.close();
        resolve(db);
      };
      req.onerror = () => reject(req.error);
    });
  }
  function idbWrite(payload) {
    return idbOpen().then(
      (db) =>
        new Promise((resolve, reject) => {
          const tx = db.transaction(IDB_STORE, "readwrite");
          tx.objectStore(IDB_STORE).put(payload, IDB_KEY);
          tx.oncomplete = () => resolve(true);
          tx.onerror = () => reject(tx.error);
          tx.onabort = () => reject(tx.error || new Error("tx aborted"));
        }),
    );
  }
  function idbRead() {
    return idbOpen().then(
      (db) =>
        new Promise((resolve, reject) => {
          const tx = db.transaction(IDB_STORE, "readonly");
          const req = tx.objectStore(IDB_STORE).get(IDB_KEY);
          req.onsuccess = () => resolve(req.result || null);
          req.onerror = () => reject(req.error);
        }),
    );
  }

  function setSaveStatus(mode, msg, autoFade) {
    const el = els.saveStatus;
    if (!el) return;
    el.hidden = false;
    el.classList.remove("saving", "error", "fade");
    if (mode === "saving") el.classList.add("saving");
    else if (mode === "error") el.classList.add("error");
    el.textContent = msg;
    if (statusTimer) {
      clearTimeout(statusTimer);
      statusTimer = null;
    }
    if (autoFade) {
      statusTimer = setTimeout(() => {
        el.classList.add("fade");
        statusTimer = setTimeout(() => {
          el.hidden = true;
          el.classList.remove("fade");
          statusTimer = null;
        }, 350);
      }, 1400);
    }
  }

  /* Single description of the document state. Used for saving, undo/redo
     snapshots and restoring, so the three cannot drift apart. Returns copies. */
  function serializeCore() {
    return {
      lists: state.lists.map((l) => ({ id: l.id, name: l.name })),
      panes: {
        left: {
          listId: state.panes.left.listId,
          search: state.panes.left.search,
        },
        right: {
          listId: state.panes.right.listId,
          search: state.panes.right.search,
        },
      },
      columns: state.columns.slice(),
      colMap: Object.assign({}, state.colMap),
      visibleColumns: Object.assign({}, state.visibleColumns),
      contacts: state.contacts.map((c) => ({
        _sid: c._sid,
        listId: c.listId,
        values: Object.assign({}, c.values),
        editedAt: c.editedAt || null,
        editKind: c.editKind || null,
      })),
      loadedName: state.loadedName,
    };
  }
  function buildPayload() {
    return Object.assign(
      { v: STORAGE_VERSION, savedAt: Date.now() },
      serializeCore(),
      {
        match: Object.assign({}, state.match),
        onlyNoNumber: state.onlyNoNumber,
        onlySelfDup: state.onlySelfDup,
      },
    );
  }

  // Set while a reset is in progress, so unload saves can't write the data back.
  let resetting = false;
  // Other open tabs of this app hold the old data in memory and would save it
  // again. A reset tells them to stop saving and reload with nothing.
  const resetChannel =
    typeof BroadcastChannel !== "undefined"
      ? new BroadcastChannel("contact_studio_reset")
      : null;
  if (resetChannel)
    resetChannel.onmessage = () => {
      resetting = true;
      window.location.reload();
    };
  /* CRITICAL: always writes localStorage synchronously. No gating. */
  function saveStateNow(opts) {
    if (resetting) return false;
    const payload = buildPayload();
    let lsOk = false,
      lsErr = null;

    if (localStorageOk) {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
        lsOk = true;
      } catch (e) {
        lsErr = e;
        if (!storageWarned) {
          storageWarned = true;
          console.warn("[Contact Studio] localStorage write failed:", e);
        }
      }
    }

    // IndexedDB backup — fire and forget
    let idbPromise = Promise.resolve(false);
    if (indexedDbOk) {
      idbPromise = idbWrite(payload).catch((e) => {
        if (!storageWarned) {
          storageWarned = true;
          console.warn("[Contact Studio] IndexedDB write failed:", e);
        }
        return false;
      });
    }

    if (opts && opts.silent !== true) {
      setSaveStatus("saving", "⏳ Saving…", false);
      idbPromise.then((idbOk) => {
        const ok = lsOk || idbOk;
        if (ok) {
          const where =
            lsOk && idbOk ? "browser + backup" : lsOk ? "browser" : "backup";
          setSaveStatus("ok", "✓ Saved · " + where, true);
        } else {
          setSaveStatus("error", "⚠ Save failed — click 💾 to retry", true);
          if (lsErr) console.error("[Contact Studio] Save error:", lsErr);
        }
      });
    }

    return lsOk;
  }

  function scheduleSave() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      saveTimer = null;
      saveStateNow({ silent: true });
    }, 200);
  }
  function flushSave() {
    if (saveTimer) {
      clearTimeout(saveTimer);
      saveTimer = null;
    }
    saveStateNow();
  }
  /* Erases everything this app stores in this browser: local and session
     storage, and the whole IndexedDB database. Resolves when done. */
  function clearStored() {
    try {
      localStorage.clear();
    } catch (e) {}
    try {
      sessionStorage.clear();
    } catch (e) {}
    if (!indexedDbOk) return Promise.resolve();
    return new Promise((resolve) => {
      const req = indexedDB.deleteDatabase(IDB_NAME);
      req.onsuccess = () => resolve();
      req.onerror = () => resolve();
    });
  }

  function resetMemory() {
    flushSave();
    const reset = () => {
      resetting = true;
      // Erase first, then set a flag so the next load also erases any copy
      // that somehow survived, and restores nothing.
      clearStored().then(() => {
        try {
          localStorage.setItem("cdf.wipe", "1");
        } catch (e) {}
        if (resetChannel) resetChannel.postMessage("reset");
        window.location.reload();
      });
    };
    if (typeof Swal !== "undefined") {
      Swal.fire({
        title: "Reset memory?",
        text: "Clears saved contacts, lists, edits, filters and column settings in this browser.",
        icon: "warning",
        showCancelButton: true,
        confirmButtonText: "Yes, reset memory",
        cancelButtonText: "Cancel",
        reverseButtons: true,
        focusCancel: true,
      }).then((r) => {
        if (r.isConfirmed) reset();
      });
    } else if (window.confirm("Reset all saved data in this browser?")) reset();
  }

  /* Validates a saved payload and returns a normalised core (or null).
     Pure: does not touch `state`. */
  function parseSavedState(data) {
    if (!data || data.v !== STORAGE_VERSION) return null;
    if (!Array.isArray(data.columns) || !Array.isArray(data.contacts))
      return null;
    const lists =
      Array.isArray(data.lists) && data.lists.length
        ? data.lists.map((l) => ({
            id: String(l.id),
            name: String(l.name || "List"),
          }))
        : [{ id: "list_1", name: "Main" }];
    const validListId = (id) => lists.some((l) => l.id === id);
    const p = data.panes || {};
    const leftId =
      p.left && validListId(p.left.listId) ? p.left.listId : lists[0].id;
    const rightId =
      p.right &&
      p.right.listId &&
      validListId(p.right.listId) &&
      p.right.listId !== leftId
        ? p.right.listId
        : (lists.find((l) => l.id !== leftId) || {}).id || null;
    const m = data.match || {};
    return {
      lists,
      panes: {
        left: { listId: leftId, search: (p.left && p.left.search) || "" },
        right: { listId: rightId, search: (p.right && p.right.search) || "" },
      },
      columns: data.columns.slice(),
      colMap: Object.assign({}, data.colMap || {}),
      visibleColumns: Object.assign({}, data.visibleColumns || {}),
      contacts: data.contacts,
      loadedName: data.loadedName || "",
      match: {
        phone: m.phone !== false,
        name: m.name !== false,
        email: m.email !== false,
      },
      onlyNoNumber: !!data.onlyNoNumber,
      onlySelfDup: !!data.onlySelfDup,
    };
  }

  /* Replaces the document state with `core` (shape from serializeCore /
     parseSavedState). All-or-nothing: if any row fails to hydrate, the
     previous state is put back and the error is rethrown. */
  function restoreCore(core) {
    const prev = {
      lists: state.lists,
      panes: state.panes,
      columns: state.columns,
      colMap: state.colMap,
      visibleColumns: state.visibleColumns,
      contacts: state.contacts,
      loadedName: state.loadedName,
    };
    try {
      state.lists = core.lists.map((l) => ({ id: l.id, name: l.name }));
      state.panes = {
        left: {
          listId: core.panes.left.listId,
          search: core.panes.left.search || "",
          selectedSids: new Set(),
        },
        right: {
          listId: core.panes.right.listId,
          search: core.panes.right.search || "",
          selectedSids: new Set(),
        },
      };
      state.columns = core.columns.slice();
      state.colMap = Object.assign({}, core.colMap);
      state.visibleColumns = Object.assign({}, core.visibleColumns);
      state.loadedName = core.loadedName || "";
      // hydrateContact reads columns/colMap/lists, so they must be set first.
      state.contacts = core.contacts.map((entry, i) =>
        hydrateContact(entry, i),
      );
      state.columns.forEach((col) => {
        if (state.visibleColumns[col] === undefined)
          state.visibleColumns[col] = true;
      });
    } catch (err) {
      Object.assign(state, prev);
      throw err;
    }
  }

  function applyLoadedData(data) {
    const core = parseSavedState(data);
    if (!core) return false;
    try {
      restoreCore(core);
    } catch (err) {
      console.warn("[Contact Studio] saved data could not be restored:", err);
      return false;
    }
    state.match = core.match;
    state.onlyNoNumber = core.onlyNoNumber;
    state.onlySelfDup = core.onlySelfDup;
    computeAllDuplicates();
    return true;
  }

  function loadStoredState() {
    // A reset was requested (flag written synchronously before reload): wipe the
    // IndexedDB backup as well and restore nothing.
    if (localStorageOk && localStorage.getItem("cdf.wipe")) {
      clearStored().then(() => {
        try {
          localStorage.removeItem("cdf.wipe");
        } catch (e) {}
      });
      return false;
    }
    let lsData = null;
    if (localStorageOk) {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) lsData = JSON.parse(raw);
      } catch (e) {
        console.warn("[Contact Studio] localStorage read failed:", e);
      }
    }
    if (lsData && applyLoadedData(lsData)) return true;

    if (indexedDbOk) {
      idbRead()
        .then((idbData) => {
          if (!idbData) return;
          if (applyLoadedData(idbData)) {
            renderAll();
            setSaveStatus("ok", "✓ Restored from backup", true);
          }
        })
        .catch((e) => {
          console.warn("[Contact Studio] IndexedDB read failed:", e);
        });
    }
    return !!(lsData && lsData.v === STORAGE_VERSION);
  }

  // This tab does not read data from other tabs: no "storage" sync listener.

  /* ========== 6. FIELD DETECTION ========== */
  function detectField(colName) {
    const raw = normalizeHeader(colName);
    if (!raw) return "other";
    const n = stripColSuffix(raw);
    if (/^e ?mail\s*\d+\s*[-–]?\s*label$/i.test(n)) return "other";
    if (
      /^(phone|mobile|cell|telephone|tel|whatsapp|fax|pager)\s*\d+\s*[-–]?\s*label$/i.test(
        n,
      )
    )
      return "phonelabel";
    if (/^labels?$/i.test(n)) return "other";
    if (/\bphonetic\b/.test(n)) return "other";
    const pairMatch = n.match(/^(.+?)\s+(\d+)\s+(label|value)$/i);
    if (pairMatch) {
      const base = pairMatch[1],
        kind = pairMatch[3].toLowerCase();
      if (/\b(phone|mobile|cell|tel|telephone|fax|whatsapp|pager)\b/.test(base))
        return kind === "label" ? "phonelabel" : "phone";
      if (/\b(e mail|email|mail)\b/.test(base))
        return kind === "label" ? "other" : "email";
    }
    if (
      /^(label|labels|phone ?label|phonelabel|phone ?type|contact ?label|type|tags?|categories)$/.test(
        n,
      )
    )
      return "phonelabel";
    if (
      /^(e ?mail|email|mail)$/.test(n) ||
      /(e ?mail|email|mail ?address|e ?mail ?address)/.test(n)
    )
      return "email";
    if (
      /^(phone|phones|mobile|cell|cellular|telephone|tel|whatsapp|fax|number|contact ?no|contact ?number|phone ?no|phone ?number|mobile ?no|mobile ?number|cell ?no|cell ?number|telephone ?no|telephone ?number|msisdn|landline|land ?line|work ?phone|home ?phone|personal ?phone|primary ?phone|secondary ?phone|other ?phone)$/.test(
        n,
      ) ||
      /\b(phone|mobile|cell|telephone|whatsapp|fax|landline|msisdn)\b/.test(
        n,
      ) ||
      /^tel(ephone)?$/.test(n)
    )
      return "phone";
    if (
      /^(full ?name|fullname|name|display ?name|contact ?name|formatted ?name|complete ?name|person|person ?name|contact|contact ?person|name ?1)$/.test(
        n,
      ) ||
      /^full ?name|^display ?name|^contact ?name|^formatted ?name|^complete ?name/.test(
        n,
      )
    )
      return "name";
    if (
      /^(first ?name|firstname|given ?name|givenname|forename|first|fname|middle ?name|middlename|middle|mname)$/.test(
        n,
      ) ||
      /first ?name|given ?name|forename|middle ?name/.test(n)
    )
      return "name";
    if (
      /^(last ?name|lastname|family ?name|familyname|surname|last|lname|sur ?name)$/.test(
        n,
      ) ||
      /last ?name|family ?name|surname/.test(n)
    )
      return "surname";
    if (
      /^(prefix|title ?prefix|salutation|honorific|suffix|nick ?name|nickname|alias)$/.test(
        n,
      ) ||
      /nick ?name|nickname|alias/.test(n)
    )
      return "other";
    if (
      /^organi[sz]ation\s+(title|department|dept|role|unit|division)$/.test(n)
    )
      return "title";
    if (
      /^(company|company ?name|organi[sz]ation|organi[sz]ation ?name|org|employer|business|business ?name|account ?name|work|workplace|firm|corp|corporation|enterprise)$/.test(
        n,
      ) ||
      /company|organi[sz]ation|employer|business|account ?name/.test(n)
    )
      return "org";
    if (
      /^(job ?title|title|position|designation|role|department|dept|profession|occupation|job|work ?title)$/.test(
        n,
      ) ||
      /job ?title|position|designation|department|profession|occupation/.test(n)
    )
      return "title";
    if (
      /^(address|address ?1|address ?2|street|city|town|state|province|county|country|zip|zip ?code|postal|postal ?code|postcode|location|region|area)$/.test(
        n,
      ) ||
      /address|street|city|state|province|country|zip|postal|location|region/.test(
        n,
      )
    )
      return "address";
    if (
      /^(note|notes|comment|comments|remark|remarks|description|memo)$/.test(
        n,
      ) ||
      /note|comment|remark|description|memo/.test(n)
    )
      return "note";
    return "other";
  }
  function detectHeader(rows) {
    if (!rows.length) return false;
    const nonEmpty = rows[0].map((c) => String(c).trim()).filter(Boolean);
    if (!nonEmpty.length) return false;
    let hits = 0;
    for (const cell of nonEmpty)
      if (HEADER_HINTS.test(cell) || detectField(cell) !== "other") hits++;
    return hits >= 1;
  }
  function dedupeHeaders(arr) {
    const seen = new Map();
    return arr.map((h, i) => {
      let name = String(h || "").trim() || "Column " + (i + 1);
      const key = name.toLowerCase();
      if (seen.has(key)) {
        const n = seen.get(key) + 1;
        seen.set(key, n);
        name += " (" + n + ")";
      } else seen.set(key, 1);
      return name;
    });
  }
  function defaultColMap() {
    const m = {};
    DEFAULT_COLUMNS.forEach((c) => {
      m[c] = detectField(c);
    });
    return m;
  }

  /* ========== 7. PARSERS ========== */
  function parseDelimited(text, delimiter) {
    const rows = [];
    let row = [],
      field = "",
      inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i++;
          } else inQuotes = false;
        } else field += ch;
      } else {
        if (ch === '"') inQuotes = true;
        else if (ch === delimiter) {
          row.push(field);
          field = "";
        } else if (ch === "\n") {
          row.push(field);
          rows.push(row);
          row = [];
          field = "";
        } else if (ch !== "\r") field += ch;
      }
    }
    row.push(field);
    rows.push(row);
    return rows;
  }
  function detectDelimiter(text) {
    const lines = text
      .split(/\r?\n/)
      .filter((l) => l.trim())
      .slice(0, 12);
    if (!lines.length) return ",";
    const sample = lines.join("\n");
    let best = ",",
      bestScore = -1;
    for (const d of [",", ";", "\t", "|"]) {
      const rows = parseDelimited(sample, d);
      if (!rows.length) continue;
      const counts = rows.map((r) => r.length);
      const max = Math.max.apply(null, counts);
      if (max < 2) continue;
      const consistent = counts.filter((c) => c === max).length / counts.length;
      const score = max * consistent;
      if (score > bestScore) {
        bestScore = score;
        best = d;
      }
    }
    return best;
  }
  function unescapeVCard(v) {
    return String(v || "")
      .replace(/\\n/gi, "\n")
      .replace(/\\,/g, ",")
      .replace(/\\;/g, ";")
      .replace(/\\\\/g, "\\")
      .trim();
  }
  function parseVCard(text) {
    const rawLines = text.split(/\r\n|\r|\n/);
    const lines = [];
    for (const line of rawLines) {
      if (/^[ \t]/.test(line) && lines.length)
        lines[lines.length - 1] += line.slice(1);
      else lines.push(line);
    }
    const cards = [];
    let cur = null;
    for (const line of lines) {
      const upper = line.toUpperCase();
      if (upper.startsWith("BEGIN:VCARD")) {
        cur = [];
        continue;
      }
      if (upper.startsWith("END:VCARD")) {
        if (cur) cards.push(cur);
        cur = null;
        continue;
      }
      if (cur) cur.push(line);
    }
    const objects = cards
      .map((cardLines) => {
        const rec = {
          fn: "",
          surname: "",
          phones: [],
          emails: [],
          org: "",
          title: "",
          note: "",
          adr: "",
        };
        for (const line of cardLines) {
          const idx = line.indexOf(":");
          if (idx === -1) continue;
          const left = line.slice(0, idx),
            value = line.slice(idx + 1);
          const prop = left.split(";")[0].trim().toUpperCase();
          switch (prop) {
            case "FN":
              if (!rec.fn) rec.fn = unescapeVCard(value);
              break;
            case "N": {
              const p = value.split(";");
              if (!rec.surname) rec.surname = unescapeVCard(p[0] || "");
              break;
            }
            case "TEL": {
              const v = unescapeVCard(value);
              if (v) {
                let label = "";
                const tm = left.match(/TYPE=([^;:]+)/i);
                if (tm)
                  for (const t of tm[1].split(/[,;]/)) {
                    const l = normalizePhoneLabel(t);
                    if (l) {
                      label = l;
                      break;
                    }
                  }
                rec.phones.push(label ? label + ": " + v : v);
              }
              break;
            }
            case "EMAIL": {
              const v = unescapeVCard(value);
              if (v) rec.emails.push(v);
              break;
            }
            case "ORG":
              if (!rec.org) rec.org = unescapeVCard(value.split(";")[0]);
              break;
            case "TITLE":
              if (!rec.title) rec.title = unescapeVCard(value);
              break;
            case "NOTE":
              if (!rec.note) rec.note = unescapeVCard(value);
              break;
            case "ADR":
              if (!rec.adr)
                rec.adr = unescapeVCard(
                  value.split(";").filter(Boolean).join(", "),
                );
              break;
          }
        }
        return {
          Name: rec.fn || "",
          Surname: rec.surname || "",
          Phone: rec.phones.join(" / "),
          Email: rec.emails.join(" / "),
          Organization: rec.org,
          Title: rec.title,
          Address: rec.adr,
          Notes: rec.note,
        };
      })
      .filter((o) => o.Name || o.Surname || o.Phone || o.Email);
    return {
      objects,
      columns: [
        "Name",
        "Surname",
        "Phone",
        "Email",
        "Organization",
        "Title",
        "Address",
        "Notes",
      ],
      colMap: {
        Name: "name",
        Surname: "surname",
        Phone: "phone",
        Email: "email",
        Organization: "org",
        Title: "title",
        Address: "address",
        Notes: "note",
      },
    };
  }
  function refineColMap(rows, columns, colMap) {
    const sample = rows.slice(0, 60);
    if (!sample.length) return colMap;
    columns.forEach((col) => {
      const current = colMap[col];
      if (current && current !== "other") return;
      if (/\bphonetic\b/i.test(col)) return;
      let filled = 0,
        emailHits = 0,
        phoneHits = 0;
      sample.forEach((row) => {
        const raw = String((row && row[col]) || "").trim();
        if (!raw) return;
        filled++;
        const arr = splitList(raw).length ? splitList(raw) : [raw];
        let isEmail = false,
          isPhone = false;
        for (const p of arr) {
          if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(p)) {
            isEmail = true;
            break;
          }
          const parsed = parseLabeledPhone(p);
          const cleanP = parsed ? parsed.number : p;
          if (normPhone(cleanP) && /^[+()\-.\s\d]+$/.test(cleanP)) {
            isPhone = true;
            break;
          }
        }
        if (isEmail) emailHits++;
        else if (isPhone) phoneHits++;
      });
      if (!filled) return;
      if (emailHits / filled >= 0.6) colMap[col] = "email";
      else if (phoneHits / filled >= 0.6) colMap[col] = "phone";
    });
    return colMap;
  }
  function parseDelimitedText(text, delimiter) {
    const rows = parseDelimited(text, delimiter)
      .map((r) => r.map(cleanCellValue))
      .filter((r) => r.some((c) => String(c).trim() !== ""));
    if (!rows.length) return { objects: [], columns: [], colMap: {} };
    let columns, dataRows;
    if (detectHeader(rows)) {
      columns = dedupeHeaders(rows[0]);
      dataRows = rows.slice(1);
    } else {
      let width = 0;
      rows.forEach((r) => {
        if (r.length > width) width = r.length;
      });
      columns = [];
      for (let i = 0; i < width; i++) columns.push("Column " + (i + 1));
      dataRows = rows;
    }
    const objects = dataRows
      .map((r) => {
        const o = {};
        columns.forEach((col, i) => {
          o[col] = r[i] == null ? "" : String(r[i]).trim();
        });
        return o;
      })
      .filter((o) => Object.keys(o).some((k) => o[k] !== ""));
    const colMap = {};
    columns.forEach((c) => {
      colMap[c] = detectField(c);
    });
    refineColMap(objects, columns, colMap);
    return { objects, columns, colMap };
  }
  function parseJSONText(text) {
    let data;
    try {
      data = JSON.parse(text);
    } catch (err) {
      throw new Error("That JSON file could not be parsed.");
    }
    let arr = null;
    if (Array.isArray(data)) arr = data;
    else if (data && Array.isArray(data.contacts)) arr = data.contacts;
    else if (data && Array.isArray(data.data)) arr = data.data;
    else if (data && typeof data === "object") arr = [data];
    if (!arr) throw new Error("JSON must be an array of contact objects.");
    const objects = arr
      .filter((o) => o && typeof o === "object" && !Array.isArray(o))
      .map((o) => {
        const rec = {};
        Object.keys(o).forEach((k) => {
          const v = o[k];
          if (v == null) rec[k] = "";
          else if (Array.isArray(v))
            rec[k] = v
              .map((x) =>
                typeof x === "object" ? JSON.stringify(x) : cleanCellValue(x),
              )
              .join(" / ");
          else if (typeof v === "object") rec[k] = JSON.stringify(v);
          else rec[k] = cleanCellValue(v);
        });
        return rec;
      });
    const columns = [];
    const seen = new Set();
    objects.forEach((o) =>
      Object.keys(o).forEach((k) => {
        if (!seen.has(k)) {
          seen.add(k);
          columns.push(k);
        }
      }),
    );
    if (!columns.length) columns = DEFAULT_COLUMNS.slice();
    const colMap = {};
    columns.forEach((c) => {
      colMap[c] = detectField(c);
    });
    refineColMap(objects, columns, colMap);
    return { objects, columns, colMap };
  }
  function parseWorkbookBuffer(buffer) {
    if (typeof XLSX === "undefined")
      throw new Error("Excel support library is still loading — please retry.");
    const wb = XLSX.read(new Uint8Array(buffer), {
      type: "array",
      cellDates: false,
      cellNF: false,
      cellText: false,
    });
    if (!wb.SheetNames || !wb.SheetNames.length)
      throw new Error("The workbook has no sheets.");
    const sheetName = wb.SheetNames[0],
      sheet = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: true,
      blankrows: false,
    });
    const rows = aoa
      .map((r) => (Array.isArray(r) ? r.map(cleanCellValue) : []))
      .filter((r) => r.some((c) => String(c).trim() !== ""));
    if (!rows.length)
      return { objects: [], columns: [], colMap: {}, sheetName };
    let columns, dataRows;
    if (detectHeader(rows)) {
      columns = dedupeHeaders(rows[0]);
      dataRows = rows.slice(1);
    } else {
      let width = 0;
      rows.forEach((r) => {
        if (r.length > width) width = r.length;
      });
      columns = [];
      for (let i = 0; i < width; i++) columns.push("Column " + (i + 1));
      dataRows = rows;
    }
    const objects = dataRows
      .map((r) => {
        const o = {};
        columns.forEach((col, i) => {
          o[col] = r[i] == null ? "" : String(r[i]).trim();
        });
        return o;
      })
      .filter((o) => Object.keys(o).some((k) => o[k] !== ""));
    const colMap = {};
    columns.forEach((c) => {
      colMap[c] = detectField(c);
    });
    refineColMap(objects, columns, colMap);
    return { objects, columns, colMap, sheetName };
  }
  function parseAnyFile(text, ext) {
    const trimmed = text.trim();
    if (!trimmed) throw new Error("The file appears to be empty.");
    if (
      ext === "vcf" ||
      ext === "vcard" ||
      /BEGIN:VCARD/i.test(trimmed.slice(0, 300))
    )
      return parseVCard(text);
    if (
      ext === "json" ||
      trimmed.charAt(0) === "[" ||
      trimmed.charAt(0) === "{"
    )
      return parseJSONText(text);
    return parseDelimitedText(
      text,
      ext === "tsv" ? "\t" : detectDelimiter(text),
    );
  }

  /* ========== 8. PHONE HELPERS ========== */
  function normalizePhoneLabel(s) {
    if (!s) return "";
    const k = String(s)
      .toLowerCase()
      .replace(/[^a-z]+/g, "");
    return k ? PHONE_LABEL_MAP[k] || "" : "";
  }
  function isKnownPhoneLabel(s) {
    if (!s) return false;
    const k = String(s)
      .toLowerCase()
      .replace(/[^a-z]+/g, "");
    return k ? Object.prototype.hasOwnProperty.call(PHONE_LABEL_MAP, k) : false;
  }
  function parseLabeledPhone(raw) {
    let s = String(raw == null ? "" : raw).trim();
    if (!s) return null;
    let label = "",
      m;
    if (
      (m = s.match(/^[\(\[]\s*([A-Za-z][A-Za-z .]{0,20}?)\s*[\)\]]\s*(.+)$/)) &&
      isKnownPhoneLabel(m[1])
    ) {
      label = normalizePhoneLabel(m[1]);
      s = m[2].trim();
    } else if (
      (m = s.match(/^([A-Za-z][A-Za-z .]{0,20}?)\s*[:=\-\u2013]\s*(.+)$/)) &&
      isKnownPhoneLabel(m[1])
    ) {
      label = normalizePhoneLabel(m[1]);
      s = m[2].trim();
    } else if (
      (m = s.match(
        /^(.+?)\s*[\(\[\{]\s*([A-Za-z][A-Za-z .]{0,20}?)\s*[\)\]\}]\s*$/,
      )) &&
      isKnownPhoneLabel(m[2])
    ) {
      label = normalizePhoneLabel(m[2]);
      s = m[1].trim();
    } else if (
      (m = s.match(/^(.+?)\s+([A-Za-z][A-Za-z .]{0,20})$/)) &&
      isKnownPhoneLabel(m[2]) &&
      /\d/.test(m[1])
    ) {
      label = normalizePhoneLabel(m[2]);
      s = m[1].trim();
    }
    // "Mobile 555-123-4567": a known label word followed by a number.
    if (
      !label &&
      (m = s.match(/^([A-Za-z][A-Za-z .]{0,20}?)\s+(\+?[\d(].*)$/)) &&
      isKnownPhoneLabel(m[1])
    ) {
      label = normalizePhoneLabel(m[1]);
      s = m[2].trim();
    }
    return { label, number: tidyPhoneNumber(s) };
  }
  // Collapse runs of spaces and strip separators left at the ends ("555-1234-").
  function tidyPhoneNumber(s) {
    return String(s || "")
      .replace(/\s+/g, " ")
      .replace(/^[\s\-–.,;:]+|[\s\-–.,;:]+$/g, "")
      .trim();
  }
  function labelFromColumnName(colName) {
    const n = normalizeHeader(colName);
    if (
      !n ||
      !/\b(phone|number|tel|mobile|cell|fax|contact|whatsapp|pager)\b/.test(n)
    )
      return "";
    for (const t of n.split(/\s+/)) {
      const l = normalizePhoneLabel(t);
      if (l) return l;
    }
    return "";
  }
  function splitPhoneListLabeled(v, fallbackLabel, rowLabel) {
    const effectiveFallback = fallbackLabel || rowLabel || "";
    return splitList(v)
      .map((part) => {
        const p = parseLabeledPhone(part);
        if (!p || !p.number || !/\d/.test(p.number)) return null;
        return { number: p.number, label: p.label || effectiveFallback || "" };
      })
      .filter(Boolean);
  }
  const titleCasePhoneLabel = (l) =>
    l
      ? String(l).charAt(0).toUpperCase() + String(l).slice(1).toLowerCase()
      : "";
  function buildRowLabelSequence(rowLabel) {
    const seq = AUTO_LABEL_SEQUENCE.slice();
    if (!rowLabel) return seq;
    const upper = String(rowLabel).toUpperCase();
    const title = titleCasePhoneLabel(upper);
    return [title].concat(seq.filter((x) => x.toUpperCase() !== upper));
  }
  function formatPhoneCellValue(v, rowLabel) {
    const raw = String(v == null ? "" : v).trim();
    if (!raw) return raw;
    const parts = splitList(raw);
    if (!parts.length) return raw;
    const tokenised = [];
    let pendingLabel = "";
    parts.forEach((part) => {
      const trimmed = String(part || "").trim();
      if (!trimmed) return;
      if (!/\d/.test(trimmed) && isKnownPhoneLabel(trimmed)) {
        pendingLabel = normalizePhoneLabel(trimmed);
        return;
      }
      const p = parseLabeledPhone(trimmed);
      if (!p) return;
      if (!/\d/.test(p.number)) {
        if (isKnownPhoneLabel(p.number))
          pendingLabel = normalizePhoneLabel(p.number);
        return;
      }
      tokenised.push({
        number: p.number,
        label: p.label || pendingLabel || "",
      });
      pendingLabel = "";
    });
    if (!tokenised.length) return raw;
    const seen = new Set(),
      deduped = [];
    tokenised.forEach((entry) => {
      const key = normPhone(entry.number) || entry.number.replace(/\D+/g, "");
      if (!key || seen.has(key)) return;
      seen.add(key);
      deduped.push(entry);
    });
    if (!deduped.length) return raw;
    const used = new Set();
    deduped.forEach((e) => {
      if (e.label) used.add(e.label);
    });
    const sequence = buildRowLabelSequence(rowLabel);
    return deduped
      .map((entry) => {
        let label = entry.label;
        if (!label) {
          for (const cand of sequence) {
            const upper = cand.toUpperCase();
            if (!used.has(upper)) {
              label = upper;
              break;
            }
          }
          if (!label) label = "OTHER";
          used.add(label);
        }
        return titleCasePhoneLabel(label) + ": " + entry.number;
      })
      .join(", ");
  }

  /* ========== 9. CONTACT MODEL ========== */
  function hydrateContact(entry, id) {
    const values = {};
    const src = (entry && entry.values) || entry || {};
    Object.keys(src).forEach((k) => {
      values[k] = cleanCellValue(src[k]);
    });
    // Phone cells may hold several numbers joined by " ::: " (Google/Outlook
    // exports). Store them as ", " so every view shows them as separate numbers.
    const formatPhonePart = (part) => {
      const p = parseLabeledPhone(part);
      if (!p || !p.number) return String(part || "").trim();
      const n = addIndiaCode(p.number);
      return p.label ? titleCasePhoneLabel(p.label) + ": " + n : n;
    };
    Object.keys(values).forEach((col) => {
      const field = state.colMap[col] || detectField(col);
      if (field === "phone" && values[col])
        values[col] = String(values[col])
          .split(/\s*:::\s*/)
          .map(formatPhonePart)
          .filter(Boolean)
          .join(", ");
    });
    const fallbackList =
      (entry && entry.listId) ||
      (state.lists[0] && state.lists[0].id) ||
      "list_1";
    const c = {
      id,
      _sid: (entry && entry._sid) || makeSid(),
      listId: fallbackList,
      values,
      name: "",
      org: "",
      phones: [],
      phoneEntries: [],
      emails: [],
      rowLabel: "",
      keys: { phone: [], email: [], name: [] },
      selfDup: false,
      selfDupPhone: false,
      selfDupEmail: false,
      groupId: null,
      searchText: "",
      editedAt: (entry && entry.editedAt) || null,
      editKind: (entry && entry.editKind) || null,
    };
    rebuildContactDerived(c, state.columns, state.colMap);
    return c;
  }
  function rebuildContactDerived(c, columns, colMap) {
    c.phones = [];
    c.phoneEntries = [];
    c.emails = [];
    c.org = "";
    c.keys = { phone: [], email: [], name: [] };
    c.rowLabel = "";
    let fullName = "",
      surname = "";
    const extraNameParts = []; // e.g. middle name; shown in the name but not used for matching
    const phoneLabelsBySlot = {};
    columns.forEach((col) => {
      const m = col.match(/^(.+?)\s+(\d+)\s*[-–]?\s*Label$/i);
      if (!m) return;
      if (!/\b(phone|mobile|cell|tel|fax|whatsapp|pager)\b/i.test(m[1])) return;
      const v = c.values[col];
      if (!v) return;
      const slot = m[2];
      const l =
        normalizePhoneLabel(String(v).trim()) ||
        (/^[A-Za-z .]{1,20}$/.test(String(v).trim())
          ? String(v).trim().toUpperCase()
          : "");
      if (l) phoneLabelsBySlot[slot] = l;
    });
    for (const col of columns) {
      const f = colMap[col] || detectField(col);
      if (f !== "phonelabel") continue;
      if (/^.+?\s+\d+\s*[-–]?\s*Label$/i.test(col)) continue;
      const v = c.values[col];
      if (!v) continue;
      const l =
        normalizePhoneLabel(String(v).trim()) ||
        (/^[A-Za-z .]{1,20}$/.test(String(v).trim())
          ? String(v).trim().toUpperCase()
          : "");
      if (l) {
        c.rowLabel = l;
        break;
      }
    }
    for (const col of columns) {
      const field = colMap[col] || detectField(col);
      const val = c.values[col];
      if (!val) continue;
      if (field === "phone") {
        const slotMatch = col.match(/^.+?\s+(\d+)\s*[-–]?\s*Value$/i);
        const columnLabel = labelFromColumnName(col);
        const slotLabel = slotMatch
          ? phoneLabelsBySlot[slotMatch[1]] || ""
          : "";
        const effectiveLabel = slotLabel || columnLabel || c.rowLabel || "";
        splitList(val).forEach((part) => {
          const p = parseLabeledPhone(part);
          if (!p || !p.number || !/\d/.test(p.number)) return;
          const label = p.label || effectiveLabel || "";
          c.phones.push(p.number);
          c.phoneEntries.push({ number: p.number, label });
        });
      } else if (field === "email") {
        splitList(val).forEach((e) => {
          if (e) c.emails.push(e);
        });
      } else if (field === "name") {
        if (!fullName) fullName = val;
        else if (
          normName(val) !== normName(fullName) &&
          !extraNameParts.some((p) => normName(p) === normName(val))
        )
          extraNameParts.push(val);
      } else if (field === "surname") {
        if (!surname) surname = val;
      } else if (field === "org") {
        if (!c.org) c.org = val;
      }
    }
    // Matching uses first + surname only, so a middle name doesn't stop a match.
    const combined = [fullName, surname].filter(Boolean).join(" ").trim();
    const displayName = [fullName]
      .concat(extraNameParts, [surname])
      .filter(Boolean)
      .join(" ")
      .trim();
    c.name = (displayName || c.emails[0] || c.phones[0] || "(no name)").trim();
    const pKeys = [],
      eKeys = [],
      nKeys = [];
    const nameKey = nameMatchKey(combined);
    if (nameKey) nKeys.push("full:" + nameKey);
    c.phones.forEach((p) => {
      const k = normPhoneKey(p);
      if (k && pKeys.indexOf(k) === -1) pKeys.push(k);
    });
    c.emails.forEach((e) => {
      const k = normEmail(e);
      if (k && eKeys.indexOf(k) === -1) eKeys.push(k);
    });
    c.keys.phone = pKeys;
    c.keys.email = eKeys;
    c.keys.name = nKeys;
    c.selfDupPhone = hasRepeatedKey(c.phones, normPhoneKey);
    c.selfDup = !!c.selfDupPhone;
    c.searchText = (
      c.name +
      " " +
      c.org +
      " " +
      c.phones.join(" ") +
      " " +
      c.emails.join(" ") +
      " " +
      Object.keys(c.values)
        .map((k) => c.values[k])
        .join(" ")
    ).toLowerCase();
  }
  function rebuildIds() {
    state.contacts.forEach((c, i) => {
      c.id = i;
    });
  }

  /* ========== 10. DUPLICATES ========== */
  function computeDuplicatesForList(listId) {
    const contacts = state.contacts;
    const inList = [];
    for (let i = 0; i < contacts.length; i++)
      if (contacts[i].listId === listId) inList.push(i);
    const parent = {};
    inList.forEach((i) => {
      parent[i] = i;
    });
    const find = (x) => {
      while (parent[x] !== x) {
        parent[x] = parent[parent[x]];
        x = parent[x];
      }
      return x;
    };
    const union = (a, b) => {
      a = find(a);
      b = find(b);
      if (a !== b) parent[b] = a;
    };
    const maps = {};
    MATCH_TYPES.forEach((t) => {
      maps[t] = new Map();
    });
    for (const i of inList) {
      const c = contacts[i];
      for (const t of MATCH_TYPES) {
        if (!state.match[t]) continue;
        const keys = c.keys[t];
        if (!keys || !keys.length) continue;
        const m = maps[t];
        for (const k of keys) {
          let arr = m.get(k);
          if (!arr) {
            arr = [];
            m.set(k, arr);
          }
          arr.push(i);
        }
      }
    }
    for (const t of MATCH_TYPES) {
      if (!state.match[t]) continue;
      for (const arr of maps[t].values())
        for (let i = 1; i < arr.length; i++) union(arr[0], arr[i]);
    }
    const byRoot = new Map();
    for (const i of inList) {
      const r = find(i);
      let arr = byRoot.get(r);
      if (!arr) {
        arr = [];
        byRoot.set(r, arr);
      }
      arr.push(i);
    }
    const groups = [];
    for (const idxs of byRoot.values()) {
      if (idxs.length < 2) continue;
      const reasons = [];
      for (const t of MATCH_TYPES) {
        if (!state.match[t]) continue;
        const seen = new Set();
        let hit = false;
        for (const i of idxs)
          for (const k of contacts[i].keys[t]) {
            if (seen.has(k)) hit = true;
            seen.add(k);
          }
        if (hit) reasons.push(t);
      }
      groups.push({
        id: listId + "_g" + groups.length,
        contacts: idxs,
        reasons: reasons.length ? reasons : ["phone"],
        color: "",
      });
    }
    groups.sort((a, b) => b.contacts.length - a.contacts.length);
    return groups;
  }
  function computeAllDuplicates() {
    state.groupsByList = new Map();
    state.groupColors = new Map();
    state.contacts.forEach((c) => {
      c.groupId = null;
    });
    state.lists.forEach((l) => {
      const groups = computeDuplicatesForList(l.id);
      state.groupsByList.set(l.id, groups);
      groups.forEach((g, i) => {
        g.color = "hsl(" + ((i * 67 + l.id.length * 17) % 360) + " 72% 46%)";
        state.groupColors.set(g.id, g.color);
        g.contacts.forEach((ci) => {
          state.contacts[ci].groupId = g.id;
        });
      });
    });
  }
  const getGroups = (listId) => state.groupsByList.get(listId) || [];

  /* ========== 11. HISTORY ========== */
  const history = { undo: [], redo: [] };
  function snapshot() {
    return serializeCore();
  }
  function pushHistory() {
    history.undo.push(snapshot());
    if (history.undo.length > HISTORY_LIMIT) history.undo.shift();
    history.redo.length = 0;
  }
  function applySnapshot(snap) {
    restoreCore(snap);
    computeAllDuplicates();
    renderAll();
    flushSave();
  }
  function undo() {
    if (!history.undo.length) return;
    history.redo.push(snapshot());
    applySnapshot(history.undo.pop());
  }
  function redo() {
    if (!history.redo.length) return;
    history.undo.push(snapshot());
    applySnapshot(history.redo.pop());
  }

  /* ========== 12. FILTERS ========== */
  function getPaneContacts(side) {
    const pane = state.panes[side];
    if (!pane.listId) return [];
    const listId = pane.listId;
    const q = (pane.search || "").trim().toLowerCase();
    return state.contacts.filter((c) => {
      if (c.listId !== listId) return false;
      if (
        state.activeGroup &&
        c.listId === state.activeGroup.listId &&
        c.groupId !== state.activeGroup.id
      )
        return false;
      if (state.onlyNoNumber && c.keys.phone.length > 0) return false;
      if (state.onlySelfDup && !c.selfDup) return false;
      if (q && c.searchText.indexOf(q) === -1) return false;
      return true;
    });
  }
  const countContactsInList = (listId) =>
    state.contacts.filter((c) => c.listId === listId).length;

  /* ========== 13. RENDERING ========== */
  function groupNumberInList(groupId) {
    for (const groups of state.groupsByList.values()) {
      const idx = groups.findIndex((g) => g.id === groupId);
      if (idx !== -1) return idx + 1;
    }
    return null;
  }
  function buildRowActions(c) {
    const wrap = document.createElement("div");
    wrap.className = "row-actions";

    // Edit (pencil) — opens the edit modal
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "act-btn";
    edit.dataset.act = "edit";
    edit.dataset.id = c._sid || String(c.id);
    edit.title = "Edit contact";
    edit.textContent = "✎";
    wrap.appendChild(edit);

    // Dedupe
    const hasRepeats = !!c.selfDup;
    const dedupe = document.createElement("button");
    dedupe.type = "button";
    dedupe.className = "act-btn dedupe" + (hasRepeats ? "" : " is-disabled");
    dedupe.dataset.act = "dedupe";
    dedupe.dataset.id = c._sid || String(c.id);
    dedupe.textContent = "1×";
    dedupe.title = hasRepeats
      ? "Remove repeated numbers"
      : "No repeated numbers";
    dedupe.disabled = !hasRepeats;
    wrap.appendChild(dedupe);

    // Delete
    const del = document.createElement("button");
    del.type = "button";
    del.className = "act-btn danger";
    del.dataset.act = "delete";
    del.dataset.id = c._sid || String(c.id);
    del.title = "Delete contact";
    del.textContent = "✕";
    wrap.appendChild(del);

    return wrap;
  }
  function renderPaneHead(side) {
    const pe = paneEls(side);
    const tr = document.createElement("tr");
    const thCheck = document.createElement("th");
    thCheck.className = "check-col";
    const allCb = document.createElement("input");
    allCb.type = "checkbox";
    allCb.className = "select-all-cb";
    allCb.title = "Select all visible";
    thCheck.appendChild(allCb);
    tr.appendChild(thCheck);
    const thIdx = document.createElement("th");
    thIdx.className = "idx";
    thIdx.textContent = "#";
    tr.appendChild(thIdx);
    const thName = document.createElement("th");
    thName.className = "name-col";
    thName.textContent = "Name";
    tr.appendChild(thName);
    const thOrg = document.createElement("th");
    thOrg.className = "org-col";
    thOrg.textContent = "Organization";
    tr.appendChild(thOrg);
    const thPhones = document.createElement("th");
    thPhones.className = "phones-col";
    thPhones.textContent = "Phones";
    tr.appendChild(thPhones);
    const thAct = document.createElement("th");
    thAct.className = "actions-th";
    tr.appendChild(thAct);
    pe.thead.innerHTML = "";
    pe.thead.appendChild(tr);
    allCb.addEventListener("change", (e) => {
      const pane = state.panes[side];
      const vis = getPaneContacts(side);
      if (e.target.checked) vis.forEach((c) => pane.selectedSids.add(c._sid));
      else vis.forEach((c) => pane.selectedSids.delete(c._sid));
      renderPane(side);
    });
  }
  function renderPane(side) {
    const pe = paneEls(side);
    const pane = state.panes[side];
    const total = pane.listId ? countContactsInList(pane.listId) : 0;
    const visible = getPaneContacts(side).length;
    pe.count.textContent =
      (visible !== total ? visible + "/" + total : String(total)) + " contacts";

    if (pe.search.value !== (pane.search || ""))
      pe.search.value = pane.search || "";

    const n = pane.selectedSids.size;
    pe.selBar.hidden = n === 0;
    pe.selCount.textContent = n + " selected";

    if (!pane.listId) {
      pe.thead.innerHTML = "";
      pe.tbody.innerHTML = "";
      if (pe.hint) {
        pe.hint.hidden = false;
        pe.hint.innerHTML =
          '<div><div class="pdh-icon">⇄</div>' +
          '<div class="pdh-title">No list shown</div>' +
          '<div class="pdh-sub">Drag rows here to create a new list, or pick one above.</div></div>';
      }
      return;
    }
    if (pe.hint) pe.hint.hidden = true;

    renderPaneHead(side);

    const vis = getPaneContacts(side);
    if (!vis.length) {
      pe.tbody.innerHTML = "";
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 6;
      td.style.cssText =
        "text-align:center;padding:26px 16px;color:#94a3b8;font-size:11.5px;border-bottom:0";
      td.textContent =
        total === 0
          ? "This list is empty — drag contacts here to add them."
          : "No contacts match your filters.";
      tr.appendChild(td);
      pe.tbody.appendChild(tr);
    } else {
      const frag = document.createDocumentFragment();
      for (const c of vis) {
        const tr = document.createElement("tr");
        tr.dataset.id = String(c.id);
        tr.dataset.sid = c._sid || "";
        tr.dataset.listId = c.listId;
        tr.draggable = true;
        if (c.selfDup) tr.classList.add("row-self-dup");
        if (c.editedAt && c.editKind === "edited")
          tr.classList.add("row-edited");
        if (pane.selectedSids.has(c._sid)) tr.classList.add("row-selected");

        const tdCheck = document.createElement("td");
        tdCheck.className = "check-col";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = pane.selectedSids.has(c._sid);
        cb.addEventListener("click", (e) => e.stopPropagation());
        cb.addEventListener("mousedown", (e) => e.stopPropagation());
        cb.addEventListener("change", (e) => {
          e.stopPropagation();
          if (cb.checked) pane.selectedSids.add(c._sid);
          else pane.selectedSids.delete(c._sid);
          tr.classList.toggle("row-selected", cb.checked);
          updateSelectionUI(side);
        });
        tdCheck.appendChild(cb);
        tr.appendChild(tdCheck);

        const tdIdx = document.createElement("td");
        tdIdx.className = "idx";
        tdIdx.textContent = String(c.id + 1);
        if (c.editedAt && c.editKind === "edited") {
          const dot = document.createElement("span");
          dot.className = "edited-dot";
          dot.title = "Edited contact";
          tdIdx.appendChild(dot);
        }
        tr.appendChild(tdIdx);

        const tdName = document.createElement("td");
        tdName.className = "name-col";
        tdName.textContent = c.name || "(no name)";
        tdName.title = c.name || "";
        tr.appendChild(tdName);

        const tdOrg = document.createElement("td");
        tdOrg.className = "org-col";
        tdOrg.textContent = c.org || "";
        tdOrg.title = c.org || "";
        tr.appendChild(tdOrg);

        const tdPhones = document.createElement("td");
        tdPhones.className = "phones-col";
        const entries =
          c.phoneEntries && c.phoneEntries.length
            ? c.phoneEntries
            : (c.phones || []).map((nn) => ({ number: nn, label: "" }));
        if (entries.length) {
          entries.slice(0, 3).forEach((e) => {
            const line = document.createElement("div");
            line.className = "phone-line";
            const lbl = e.label ? titleCasePhoneLabel(e.label) : "";
            line.textContent = lbl ? lbl + " – " + e.number : e.number;
            tdPhones.appendChild(line);
          });
          if (entries.length > 3) {
            const more = document.createElement("div");
            more.className = "phone-line muted";
            more.textContent = "+" + (entries.length - 3) + " more";
            tdPhones.appendChild(more);
          }
        } else {
          const line = document.createElement("div");
          line.className = "phone-line muted";
          line.textContent = "—";
          tdPhones.appendChild(line);
        }
        tr.appendChild(tdPhones);

        const tdActions = document.createElement("td");
        tdActions.className = "actions-cell";
        tdActions.appendChild(buildRowActions(c));
        tr.appendChild(tdActions);

        // Double-click row → edit
        tr.addEventListener("dblclick", (e) => {
          if (e.target.closest && e.target.closest("input[type=checkbox]"))
            return;
          if (e.target.closest && e.target.closest("button[data-act]")) return;
          openContactModal(c._sid);
        });


        frag.appendChild(tr);
      }
      pe.tbody.innerHTML = "";
      pe.tbody.appendChild(frag);
    }

    const allCb = pe.thead.querySelector(".select-all-cb");
    if (allCb) {
      const visibleAll = getPaneContacts(side);
      allCb.checked =
        visibleAll.length > 0 &&
        visibleAll.every((c) => pane.selectedSids.has(c._sid));
      allCb.indeterminate =
        !allCb.checked && visibleAll.some((c) => pane.selectedSids.has(c._sid));
    }
  }
  function updateSelectionUI(side) {
    const pe = paneEls(side);
    const n = state.panes[side].selectedSids.size;
    pe.selBar.hidden = n === 0;
    pe.selCount.textContent = n + " selected";
  }
  function renderDupesDrawer() {
    const drawer = els.dupesDrawer;
    if (!drawer || drawer.hidden) return;
    const list = els.dupesList;
    list.innerHTML = "";
    if (state.activeGroup) {
      const clear = document.createElement("button");
      clear.type = "button";
      clear.className = "group-clear-btn";
      clear.textContent = "✕ Show all contacts";
      clear.addEventListener("click", () => {
        state.activeGroup = null;
        renderAll();
      });
      list.appendChild(clear);
    }
    const curListId = state.panes.left.listId;
    if (!curListId) {
      list.innerHTML =
        '<div class="sb-empty"><span class="big">⌕</span><div><b>No list selected</b></div></div>';
      return;
    }
    const groups = getGroups(curListId);
    if (!groups.length) {
      list.innerHTML =
        '<div class="sb-empty"><span class="big">🎉</span><div><b>No duplicate groups</b><small>in list “' +
        escapeHtml(listName(curListId)) +
        "”</small></div></div>";
      return;
    }
    const frag = document.createDocumentFragment();
    groups.forEach((g, gi) => {
      const first = state.contacts[g.contacts[0]];
      const extra = g.contacts.length - 1;
      const card = document.createElement("div");
      card.className =
        "group-card" +
        (state.activeGroup && state.activeGroup.id === g.id ? " active" : "");
      card.style.setProperty("--g", g.color);
      const head = document.createElement("div");
      head.className = "gc-head";
      const dot = document.createElement("span");
      dot.className = "gc-dot";
      head.appendChild(dot);
      const title = document.createElement("span");
      title.className = "gc-title";
      title.textContent = "Group " + (gi + 1);
      head.appendChild(title);
      const count = document.createElement("span");
      count.className = "gc-count";
      count.textContent = g.contacts.length;
      head.appendChild(count);
      card.appendChild(head);
      const body = document.createElement("div");
      body.className = "gc-body";
      const nm = document.createElement("span");
      nm.className = "gc-name";
      nm.textContent = first && first.name ? first.name : "(no name)";
      nm.title = nm.textContent;
      body.appendChild(nm);
      if (extra > 0) {
        const more = document.createElement("span");
        more.className = "gc-more";
        more.textContent = "+" + extra + " more";
        body.appendChild(more);
      }
      if (g.reasons.length) {
        const badges = document.createElement("span");
        badges.className = "badges";
        g.reasons.forEach((r) => {
          const b = document.createElement("span");
          b.className = "badge " + r;
          b.textContent = REASON_LABEL[r] || r;
          badges.appendChild(b);
        });
        body.appendChild(badges);
      }
      card.appendChild(body);
      // Click toggles this group: show only its contacts in the list, or show all again.
      card.addEventListener("click", () => {
        const on = state.activeGroup && state.activeGroup.id === g.id;
        state.activeGroup = on ? null : { id: g.id, listId: curListId };
        renderAll();
      });
      frag.appendChild(card);
    });
    list.appendChild(frag);
  }
  function renderAll() {
    // Drop the group filter if that group no longer exists after an edit.
    if (
      state.activeGroup &&
      !getGroups(state.activeGroup.listId).some(
        (g) => g.id === state.activeGroup.id,
      )
    )
      state.activeGroup = null;
    renderPane("left");
    renderDupesDrawer();
  }

  /* ========== 15. MODALS ========== */

  function createCollapsibleSection(title, icon, count, defaultExpanded) {
    const section = document.createElement("div");
    section.className =
      "details-section" + (defaultExpanded ? " expanded" : "");
    const head = document.createElement("div");
    head.className = "details-section-title";
    head.setAttribute("role", "button");
    head.setAttribute("tabindex", "0");
    head.innerHTML =
      '<span class="details-section-head">' +
      '<span class="details-section-icon">' +
      icon +
      "</span>" +
      "<span>" +
      escapeHtml(title) +
      "</span>" +
      (count != null
        ? '<span class="details-section-count">' + count + "</span>"
        : "") +
      "</span>" +
      '<span class="details-section-caret">▶</span>';
    const body = document.createElement("div");
    body.className = "details-section-body";
    body.hidden = !defaultExpanded;
    const toggle = () => {
      const isExp = section.classList.toggle("expanded");
      body.hidden = !isExp;
    };
    head.addEventListener("click", toggle);
    head.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        toggle();
      }
    });
    section.appendChild(head);
    section.appendChild(body);
    return { section, head, body };
  }

  function openContactDetails(contactId) {
    const contact = state.contacts.find((c) =>
      c._sid ? c._sid === contactId : c.id === Number(contactId),
    );
    if (!contact) return;
    state.detailsContactSid = contact._sid || String(contact.id);
    const group = contact.groupId
      ? (() => {
          for (const gs of state.groupsByList.values()) {
            const g = gs.find((x) => x.id === contact.groupId);
            if (g) return g;
          }
          return null;
        })()
      : null;
    const duplicateReasons = group ? group.reasons || [] : [];
    const initials =
      String(contact.name || "?")
        .split(/\s+/)
        .filter(Boolean)
        .slice(0, 2)
        .map((x) => x.charAt(0).toUpperCase())
        .join("") || "?";
    els.detailsModalTitle.textContent = contact.name || "Contact details";
    els.detailsBody.innerHTML = "";

    const head = document.createElement("div");
    head.className = "details-profile";
    head.innerHTML =
      '<div class="details-avatar">' +
      escapeHtml(initials) +
      "</div>" +
      '<div class="details-profile-main">' +
      '<div class="details-name">' +
      escapeHtml(contact.name || "Unnamed contact") +
      "</div>" +
      '<div class="details-meta">Record #' +
      (contact.id + 1) +
      (state.loadedName ? " · " + escapeHtml(state.loadedName) : "") +
      (contact.org ? " · " + escapeHtml(contact.org) : "") +
      "</div></div>";
    els.detailsBody.appendChild(head);

    const chips = document.createElement("div");
    chips.className = "details-chips";
    const chipData = [];
    if (group)
      chipData.push({
        text:
          "G" +
          groupNumberInList(group.id) +
          " · " +
          duplicateReasons.map((x) => REASON_LABEL[x] || x).join(" + "),
        cls: "dup",
      });
    else chipData.push({ text: "No duplicate match", cls: "muted" });
    if (contact.selfDup)
      chipData.push({ text: "Self-duplicate · phone", cls: "self" });
    // Organization shown as a badge.
    if (contact.org)
      chipData.push({ text: "Org · " + contact.org, cls: "plain" });
    chipData.forEach((item) => {
      const chip = document.createElement("span");
      chip.className = "details-chip " + item.cls;
      chip.textContent = item.text;
      chips.appendChild(chip);
    });
    els.detailsBody.appendChild(chips);

    const fieldCount = (() => {
      let n = 0;
      state.columns.forEach((col) => {
        const v = contact.values ? contact.values[col] : "";
        if (v == null || String(v).trim() === "") return;
        n++;
      });
      Object.keys(contact.values || {}).forEach((col) => {
        if (state.columns.includes(col)) return;
        const v = contact.values[col];
        if (v == null || String(v).trim() === "") return;
        n++;
      });
      return n;
    })();
    const fieldsSection = createCollapsibleSection(
      "All contact information",
      "☰",
      fieldCount + " field" + (fieldCount === 1 ? "" : "s"),
      true,
    );
    const rows = document.createElement("div");
    rows.className = "details-fields";
    const shown = new Set();
    state.columns.forEach((col) => {
      const raw = contact.values ? contact.values[col] : "";
      const value = raw == null ? "" : String(raw).trim();
      if (!value) return;
      shown.add(col);
      const row = document.createElement("div");
      row.className = "detail-row";
      row.innerHTML =
        '<div class="detail-label">' +
        escapeHtml(col) +
        '</div><div class="detail-value">' +
        escapeHtml(value) +
        "</div>";
      rows.appendChild(row);
    });
    Object.keys(contact.values || {}).forEach((col) => {
      if (shown.has(col)) return;
      const value = contact.values[col];
      if (value == null || String(value).trim() === "") return;
      const row = document.createElement("div");
      row.className = "detail-row";
      row.innerHTML =
        '<div class="detail-label">' +
        escapeHtml(col) +
        '</div><div class="detail-value">' +
        escapeHtml(String(value)) +
        "</div>";
      rows.appendChild(row);
    });
    if (!rows.childNodes.length)
      rows.innerHTML =
        '<div class="details-empty">No contact fields available.</div>';
    fieldsSection.body.appendChild(rows);
    els.detailsBody.appendChild(fieldsSection.section);

    if (group) {
      const grpSection = createCollapsibleSection(
        "Duplicate group",
        "◆",
        group.contacts.length +
          " contact" +
          (group.contacts.length === 1 ? "" : "s"),
        true,
      );
      const members = document.createElement("div");
      members.className = "details-members";
      group.contacts.slice(0, 8).forEach((idx) => {
        const member = state.contacts[idx];
        if (!member) return;
        const item = document.createElement("div");
        item.className =
          "details-member" + (member._sid === contact._sid ? " current" : "");
        item.innerHTML =
          "<span>" +
          escapeHtml(member.name || "Unnamed contact") +
          "</span><small>#" +
          (member.id + 1) +
          "</small>";
        members.appendChild(item);
      });
      if (group.contacts.length > 8) {
        const more = document.createElement("div");
        more.className = "details-more";
        more.textContent =
          "+ " + (group.contacts.length - 8) + " more contacts";
        members.appendChild(more);
      }
      grpSection.body.appendChild(members);
      els.detailsBody.appendChild(grpSection.section);
    }

    if (els.detailsModalDedupeBtn)
      els.detailsModalDedupeBtn.disabled = !contact.selfDup;
    els.detailsModal.hidden = false;
  }
  function closeContactDetails() {
    els.detailsModal.hidden = true;
    els.detailsBody.innerHTML = "";
    state.detailsContactSid = null;
  }

  /* Custom dropdown for the phone label. The native <select> stays in the form
     (hidden) so saving can still read its .value. */
  let openLabelMenu = null;
  function closeLabelMenu() {
    if (!openLabelMenu) return;
    const { wrap, menu, trigger } = openLabelMenu;
    menu.hidden = true;
    menu.removeAttribute("style");
    wrap.appendChild(menu);
    trigger.setAttribute("aria-expanded", "false");
    openLabelMenu = null;
  }
  /* The open menu is moved to <body> and positioned against its trigger, so
     the modal's scrolling body and footer can't clip it. It opens upward when
     there is no room below. */
  function openLabelMenuAt(trigger, menu) {
    document.body.appendChild(menu);
    menu.hidden = false;
    menu.style.position = "fixed";
    menu.style.zIndex = "1000";
    const r = trigger.getBoundingClientRect();
    menu.style.left = r.left + "px";
    menu.style.width = r.width + "px";
    menu.style.top = r.bottom + 4 + "px";
    const h = menu.offsetHeight;
    if (r.bottom + 4 + h > window.innerHeight - 8 && r.top - 4 - h > 8)
      menu.style.top = r.top - 4 - h + "px";
    trigger.setAttribute("aria-expanded", "true");
  }
  document.addEventListener("click", (e) => {
    if (openLabelMenu && !openLabelMenu.wrap.contains(e.target))
      closeLabelMenu();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeLabelMenu();
  });
  // Close on any scroll so the menu never sits away from its trigger.
  document.addEventListener("scroll", closeLabelMenu, true);
  window.addEventListener("resize", closeLabelMenu);
  function buildLabelPicker(select) {
    const wrap = document.createElement("div");
    wrap.className = "label-picker";
    const trigger = document.createElement("button");
    trigger.type = "button";
    trigger.className = "label-picker-trigger";
    trigger.setAttribute("aria-haspopup", "listbox");
    trigger.setAttribute("aria-expanded", "false");
    const menu = document.createElement("div");
    menu.className = "label-picker-menu";
    menu.setAttribute("role", "listbox");
    menu.hidden = true;
    const sync = () => {
      const opt = select.options[select.selectedIndex];
      trigger.textContent = opt ? opt.textContent : "No label";
      menu.querySelectorAll(".label-picker-option").forEach((b) => {
        const on = b.dataset.value === select.value;
        b.classList.toggle("selected", on);
        b.setAttribute("aria-selected", String(on));
      });
    };
    Array.from(select.options).forEach((o) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "label-picker-option";
      b.setAttribute("role", "option");
      b.dataset.value = o.value;
      b.textContent = o.textContent;
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        select.value = o.value;
        sync();
        closeLabelMenu();
      });
      menu.appendChild(b);
    });
    trigger.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = openLabelMenu && openLabelMenu.menu === menu;
      closeLabelMenu();
      if (isOpen) return;
      openLabelMenuAt(trigger, menu);
      openLabelMenu = { wrap, menu, trigger };
    });
    select.hidden = true;
    select.parentNode.insertBefore(wrap, select);
    wrap.append(trigger, menu);
    select.addEventListener("change", sync);
    sync();
  }
  function openContactModal(contactId, defaultListId) {
    const isEdit = contactId != null && contactId !== "";
    let columns = state.columns.slice(),
      colMap = state.colMap;
    if (!columns.length) {
      columns = DEFAULT_COLUMNS.slice();
      colMap = defaultColMap();
      state.columns = columns.slice();
      state.colMap = colMap;
    }
    const hasName = columns.some(
      (c) => (colMap[c] || detectField(c)) === "name",
    );
    const hasSurname = columns.some(
      (c) => (colMap[c] || detectField(c)) === "surname",
    );
    if (!hasName) {
      let nn = "Name";
      while (columns.includes(nn)) nn += " ";
      columns.push(nn);
      colMap[nn] = "name";
      state.columns.push(nn);
      state.colMap[nn] = "name";
    }
    if (!hasSurname) {
      let nn = "Surname";
      while (columns.includes(nn)) nn += " ";
      columns.push(nn);
      colMap[nn] = "surname";
      state.columns.push(nn);
      state.colMap[nn] = "surname";
    }
    const editing = isEdit
      ? state.contacts.find((c) =>
          c._sid ? c._sid === contactId : c.id === Number(contactId),
        )
      : null;
    state.editingId = editing ? editing._sid || editing.id : null;
    state.editingListId =
      defaultListId ||
      (editing && editing.listId) ||
      state.panes.left.listId ||
      state.lists[0].id;
    els.modalTitle.textContent = isEdit
      ? "✎ Edit contact"
      : "＋ Add new contact";
    const body = els.modalBody;
    body.innerHTML = "";
    const help = document.createElement("div");
    help.className = "contact-form-help";
    help.innerHTML =
      "Only <b>Name</b> and <b>Surname</b> are captured. Add multiple phone numbers with labels below.";
    body.appendChild(help);
    const seed = editing ? Object.assign({}, editing.values) : {};
    const phoneColumns = columns.filter(
      (col) => (colMap[col] || detectField(col)) === "phone",
    );
    const phoneLabelColumns = columns.filter(
      (col) => (colMap[col] || detectField(col)) === "phonelabel",
    );
    const primaryPhoneCol = phoneColumns[0] || "Phone";
    const primaryLabelCol = phoneLabelColumns[0] || "";
    const phoneSet = new Set(phoneColumns),
      labelSet = new Set(phoneLabelColumns);
    const nameCols = columns.filter(
      (c) => (colMap[c] || detectField(c)) === "name",
    );
    const surnameCols = columns.filter(
      (c) => (colMap[c] || detectField(c)) === "surname",
    );
    const primaryNameCol = nameCols[0];
    const primarySurnameCol = surnameCols[0];
    const nameSurnameSet = new Set([...nameCols, ...surnameCols]);
    const frag = document.createDocumentFragment();
    const renderField = (col, fieldType, labelText, placeholder) => {
      const row = document.createElement("div");
      row.className = "form-row";
      const lbl = document.createElement("label");
      lbl.textContent = labelText;
      const type = document.createElement("span");
      type.className = "field-type " + fieldType;
      type.textContent = fieldType;
      lbl.appendChild(type);
      row.appendChild(lbl);
      const inp = document.createElement("input");
      inp.type = "text";
      inp.dataset.col = col;
      inp.autocomplete = "off";
      inp.spellcheck = false;
      inp.value = seed[col] || "";
      if (placeholder) inp.placeholder = placeholder;
      row.appendChild(inp);
      frag.appendChild(row);
    };
    renderField(primaryNameCol, "name", "Name", "Given / full name");
    renderField(primarySurnameCol, "surname", "Surname", "Family name");
    const HIDDEN =
      /^(e ?mail ?\d+ ?(value|type|label)|organi[sz]ation ?(department|dept|title|role|unit|division))$/i;
    const allowed = ["org", "title", "email"];
    let orgRendered = false;
    columns.forEach((col) => {
      if (phoneSet.has(col) || labelSet.has(col)) return;
      if (nameSurnameSet.has(col)) return;
      if (HIDDEN.test(normalizeHeader(col))) return;
      const field = colMap[col] || detectField(col);
      if (!allowed.includes(field)) return;
      if (field === "org") {
        if (orgRendered) return;
        orgRendered = true;
      }
      renderField(col, field, col, "");
    });
    const section = document.createElement("div");
    section.className = "phone-editor-section";
    const head = document.createElement("div");
    head.className = "phone-editor-head";
    head.innerHTML =
      "<div><strong>Phone numbers</strong><small>Multiple numbers are supported</small></div>" +
      '<button type="button" class="phone-add-btn">＋ Add number</button>';
    section.appendChild(head);
    const list = document.createElement("div");
    list.className = "phone-editor-list";
    section.appendChild(list);
    const entries = [];
    phoneColumns.forEach((col) => {
      if (!seed[col]) return;
      splitPhoneListLabeled(
        seed[col],
        labelFromColumnName(col),
        seed[primaryLabelCol] || "",
      ).forEach((e) =>
        entries.push({ number: e.number, label: e.label || "" }),
      );
    });
    if (
      !entries.length &&
      editing &&
      editing.phoneEntries &&
      editing.phoneEntries.length
    ) {
      editing.phoneEntries.forEach((e) =>
        entries.push({ number: e.number, label: e.label || "" }),
      );
    }
    if (!entries.length)
      entries.push({ number: "", label: seed[primaryLabelCol] || "" });
    function addPhoneRow(entry) {
      entry = entry || { number: "", label: "" };
      const row = document.createElement("div");
      row.className = "phone-editor-row";
      row.innerHTML =
        '<select class="phone-label-input">' +
        '<option value="">No label</option>' +
        '<option value="MAIN">Main</option><option value="WORK">Work</option>' +
        '<option value="HOME">Home</option><option value="CELL">Cell</option>' +
        '<option value="WHATSAPP">WhatsApp</option><option value="OTHER">Other</option>' +
        "</select>" +
        '<input class="phone-number-input" type="text" placeholder="+91 12345 12345">' +
        '<button type="button" class="phone-remove-btn" title="Remove number">×</button>';
      row.querySelector(".phone-label-input").value =
        normalizePhoneLabel(entry.label) || "";
      row.querySelector(".phone-number-input").value = entry.number || "";
      buildLabelPicker(row.querySelector(".phone-label-input"));
      row.querySelector(".phone-remove-btn").addEventListener("click", () => {
        if (list.children.length === 1) {
          row.querySelector(".phone-label-input").value = "";
          row
            .querySelector(".phone-label-input")
            .dispatchEvent(new Event("change"));
          row.querySelector(".phone-number-input").value = "";
        } else row.remove();
      });
      list.appendChild(row);
    }
    entries.forEach(addPhoneRow);
    head.querySelector(".phone-add-btn").addEventListener("click", () => {
      addPhoneRow();
      const last = list.lastElementChild;
      if (last) {
        const inp = last.querySelector(".phone-number-input");
        if (inp) inp.focus();
      }
    });
    frag.appendChild(section);
    body.appendChild(frag);
    body.dataset.primaryPhoneCol = primaryPhoneCol;
    body.dataset.primaryLabelCol = primaryLabelCol;
    body.dataset.phoneColumns = JSON.stringify(phoneColumns);
    els.contactModal.hidden = false;
    setTimeout(() => {
      const inp = body.querySelector(".phone-number-input");
      if (inp) inp.focus();
    }, 40);
  }
  function closeContactModal() {
    els.contactModal.hidden = true;
    state.editingId = null;
    els.modalBody.innerHTML = "";
  }
  function saveContact(e) {
    e.preventDefault();
    if (!state.columns.length) {
      state.columns = DEFAULT_COLUMNS.slice();
      state.colMap = defaultColMap();
      state.columns.forEach((col) => {
        if (state.visibleColumns[col] === undefined)
          state.visibleColumns[col] = true;
      });
    }
    const values = {};
    els.modalBody.querySelectorAll("input[data-col]").forEach((inp) => {
      values[inp.dataset.col] = cleanCellValue(inp.value);
    });
    const phoneRows = Array.from(
      els.modalBody.querySelectorAll(".phone-editor-row"),
    );
    const primaryPhoneCol = els.modalBody.dataset.primaryPhoneCol || "Phone";
    const primaryLabelCol = els.modalBody.dataset.primaryLabelCol || "";
    let phoneColumns = [];
    try {
      phoneColumns = JSON.parse(els.modalBody.dataset.phoneColumns || "[]");
    } catch (_) {}
    if (phoneRows.length) {
      const numbers = [],
        labels = [];
      phoneRows.forEach((row) => {
        const raw = cleanCellValue(
          row.querySelector(".phone-number-input")
            ? row.querySelector(".phone-number-input").value
            : "",
        );
        const label = normalizePhoneLabel(
          row.querySelector(".phone-label-input")
            ? row.querySelector(".phone-label-input").value
            : "",
        );
        if (!raw) return;
        splitList(raw).forEach((part) => {
          if (part) {
            const num = addIndiaCode(part);
            numbers.push(label ? label + ": " + num : num);
            labels.push(label);
          }
        });
      });
      if (phoneColumns.length) {
        phoneColumns.forEach((col, i) => {
          values[col] = i === 0 ? numbers.join(" / ") : "";
        });
      } else values[primaryPhoneCol] = numbers.join(" / ");
      if (primaryLabelCol && !values[primaryLabelCol])
        values[primaryLabelCol] = labels[0] || "";
    }
    const isEdit = state.editingId != null;
    const now = Date.now();
    pushHistory();
    if (isEdit) {
      const target = state.contacts.find((x) =>
        x._sid ? x._sid === state.editingId : x.id === Number(state.editingId),
      );
      if (!target) {
        closeContactModal();
        return;
      }
      target.values = Object.assign({}, values);
      target.editedAt = now;
      if (target.editKind !== "added") target.editKind = "edited";
      rebuildContactDerived(target, state.columns, state.colMap);
    } else {
      const listId =
        state.editingListId || state.panes.left.listId || state.lists[0].id;
      const c = hydrateContact(
        {
          values: Object.assign({}, values),
          editedAt: now,
          editKind: "added",
          listId,
        },
        state.contacts.length,
      );
      state.contacts.push(c);
    }
    rebuildIds();
    closeContactModal();
    computeAllDuplicates();
    renderAll();
    flushSave();
  }
  function autoLabelRow(sid) {
    const c = state.contacts.find((x) => x._sid === sid);
    if (!c) return;
    let changed = 0;
    state.columns.forEach((col) => {
      const field = state.colMap[col] || detectField(col);
      if (field !== "phone") return;
      const raw = c.values[col];
      if (raw == null) return;
      const s = String(raw).trim();
      if (!s) return;
      const next = formatPhoneCellValue(s, c.rowLabel);
      if (next !== s) {
        c.values[col] = next;
        changed++;
      }
    });
    if (!changed) {
      toast("Already labelled", "Nothing to change.", "info");
      return;
    }
    pushHistory();
    c.editedAt = Date.now();
    if (c.editKind !== "added") c.editKind = "edited";
    rebuildContactDerived(c, state.columns, state.colMap);
    computeAllDuplicates();
    renderAll();
    flushSave();
    toast("Phone list cleaned", changed + " field(s) updated.", "success");
  }
  function dedupeRow(sid) {
    const c = state.contacts.find((x) => x._sid === sid);
    if (!c) return;
    if (!c.selfDup) {
      toast("Nothing to clean", "No repeated phone numbers.", "info");
      return;
    }
    const seen = new Set();
    let changed = 0;
    state.columns.forEach((col) => {
      if ((state.colMap[col] || detectField(col)) !== "phone") return;
      const raw = c.values[col];
      if (raw == null) return;
      const s = String(raw).trim();
      if (!s) return;
      const parts = splitList(s);
      if (!parts.length) return;
      const rebuilt = [];
      parts.forEach((part) => {
        const trimmed = String(part || "").trim();
        if (!trimmed) return;
        const p = parseLabeledPhone(trimmed);
        if (!p || !p.number || !/\d/.test(p.number)) {
          rebuilt.push(trimmed);
          return;
        }
        const key = normPhone(p.number) || p.number.replace(/\D+/g, "");
        if (!key) {
          rebuilt.push(trimmed);
          return;
        }
        if (seen.has(key)) return;
        seen.add(key);
        rebuilt.push(
          p.label ? titleCasePhoneLabel(p.label) + ": " + p.number : p.number,
        );
      });
      const next = rebuilt.join(", ");
      if (next !== s) {
        c.values[col] = next;
        changed++;
      }
    });
    if (!changed) {
      toast("Nothing to clean", "No repeated phone numbers.", "info");
      return;
    }
    pushHistory();
    c.editedAt = Date.now();
    if (c.editKind !== "added") c.editKind = "edited";
    rebuildContactDerived(c, state.columns, state.colMap);
    computeAllDuplicates();
    renderAll();
    flushSave();
    toast(
      "Repeated numbers removed",
      "Kept one copy of each phone number.",
      "success",
    );
  }

  /* ========== 16. FILE I/O + EXPORT ========== */
  function confirmAction(opts) {
    if (typeof Swal === "undefined")
      return Promise.resolve(
        window.confirm(opts.fallback || opts.title || "Are you sure?"),
      );
    return Swal.fire({
      title: opts.title,
      html: opts.html,
      icon: opts.icon || "warning",
      showCancelButton: true,
      confirmButtonText: opts.confirmText || "Confirm",
      cancelButtonText: opts.cancelText || "Cancel",
      reverseButtons: true,
      focusCancel: !!opts.focusCancel,
    }).then((r) => !!(r && r.isConfirmed));
  }
  function toast(title, text, icon) {
    if (typeof Swal === "undefined") {
      if (text) console.log(title + " — " + text);
      return;
    }
    Swal.fire({
      toast: true,
      position: "bottom-end",
      icon: icon || "info",
      title,
      text: text || "",
      showConfirmButton: false,
      timer: 2600,
      timerProgressBar: true,
    });
  }
  function downloadFile(filename, content, mime) {
    const blob = new Blob([content], {
      type: mime || "text/plain;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      if (a.parentNode) a.parentNode.removeChild(a);
    }, 0);
  }
  const escVCard = (s) =>
    String(s == null ? "" : s)
      .replace(/\\/g, "\\\\")
      .replace(/\r?\n/g, "\\n")
      .replace(/,/g, "\\,")
      .replace(/;/g, "\\;");
  function collectField(c, field) {
    const out = [];
    state.columns.forEach((col) => {
      if (state.colMap[col] !== field) return;
      const v = c.values[col];
      if (v == null) return;
      const s = String(v).trim();
      if (!s) return;
      out.push(s);
    });
    return out;
  }
  const vcfTypeForLabel = (label) =>
    !label
      ? "CELL,VOICE"
      : VCF_TYPE_MAP[String(label).toUpperCase()] || "VOICE";
  function contactToVCard(c) {
    const names = collectField(c, "name"),
      surnames = collectField(c, "surname");
    const phoneEntries =
      c.phoneEntries && c.phoneEntries.length
        ? c.phoneEntries.slice()
        : (() => {
            const out = [];
            state.columns.forEach((col) => {
              if (state.colMap[col] !== "phone") return;
              const v = c.values[col];
              if (v == null) return;
              const s = String(v).trim();
              if (!s) return;
              const fallback = labelFromColumnName(col);
              splitPhoneListLabeled(s, fallback, c.rowLabel).forEach((e) => {
                if (e && e.number) out.push(e);
              });
            });
            return out;
          })();
    const emails = [];
    collectField(c, "email").forEach((v) =>
      splitList(v).forEach((m) => emails.push(m)),
    );
    const orgs = collectField(c, "org"),
      titles = collectField(c, "title");
    const addrs = collectField(c, "address"),
      notes = collectField(c, "note");
    const given = names[0] || "",
      family = surnames[0] || "";
    const full = [given, family].filter(Boolean).join(" ") || c.name || "";
    const lines = ["BEGIN:VCARD", "VERSION:3.0"];
    lines.push("FN:" + escVCard(full));
    lines.push("N:" + escVCard(family) + ";" + escVCard(given) + ";;;");
    phoneEntries.forEach((e) =>
      lines.push(
        "TEL;TYPE=" + vcfTypeForLabel(e.label) + ":" + escVCard(e.number),
      ),
    );
    emails.forEach((m) => lines.push("EMAIL;TYPE=INTERNET:" + escVCard(m)));
    if (orgs[0]) lines.push("ORG:" + escVCard(orgs[0]));
    if (titles[0]) lines.push("TITLE:" + escVCard(titles[0]));
    addrs.forEach((a) => lines.push("ADR;TYPE=HOME:;;" + escVCard(a) + ";;;;"));
    notes.forEach((n) => lines.push("NOTE:" + escVCard(n)));
    lines.push("END:VCARD");
    return lines.join("\r\n");
  }
  const exportContactsVcf = (contacts) =>
    contacts.map(contactToVCard).join("\r\n") + "\r\n";
  function exportContactsDelimited(contacts, delimiter) {
    const cols = state.columns.slice();
    const rows = [["#", "List", "Grp"].concat(cols)];
    contacts.forEach((c) => {
      const group = c.groupId ? "G" + (groupNumberInList(c.groupId) || "") : "";
      rows.push(
        [String(c.id + 1), listName(c.listId), group].concat(
          cols.map((col) => {
            const v = c.values ? c.values[col] : "";
            return v == null ? "" : String(v);
          }),
        ),
      );
    });
    return (
      rows
        .map((row) =>
          row
            .map((v) => {
              const value = String(v == null ? "" : v);
              if (delimiter === "\t")
                return value.replace(/\r?\n/g, " ").replace(/\t/g, " ");
              return /[",\r\n]/.test(value)
                ? '"' + value.replace(/"/g, '""') + '"'
                : value;
            })
            .join(delimiter),
        )
        .join("\r\n") + "\r\n"
    );
  }
  function exportContactsJson(contacts) {
    const data = contacts.map((c) => ({
      id: c.id + 1,
      list: listName(c.listId),
      name: c.name || "",
      org: c.org || "",
      phones: Array.isArray(c.phones) ? c.phones.slice() : [],
      phoneEntries: (c.phoneEntries || []).map((e) => ({
        number: e.number,
        label: e.label || "",
      })),
      emails: Array.isArray(c.emails) ? c.emails.slice() : [],
      selfDuplicate: !!c.selfDup,
      group: c.groupId ? "G" + (groupNumberInList(c.groupId) || "") : null,
      editedAt: c.editedAt || null,
      editKind: c.editKind || null,
      values: Object.assign({}, c.values || {}),
    }));
    return (
      JSON.stringify(
        {
          exportedAt: new Date().toISOString(),
          sourceFile: state.loadedName || null,
          count: data.length,
          contacts: data,
        },
        null,
        2,
      ) + "\n"
    );
  }
  function exportContacts(contacts, format) {
    if (!contacts || !contacts.length) {
      toast("Nothing to export", "This list is empty.", "info");
      return;
    }
    let filename = "contacts_" + contacts.length,
      content = "",
      mime = "text/plain;charset=utf-8";
    if (format === "vcf") {
      filename += ".vcf";
      content = exportContactsVcf(contacts);
      mime = "text/vcard;charset=utf-8";
    } else if (format === "json") {
      filename += ".json";
      content = exportContactsJson(contacts);
      mime = "application/json;charset=utf-8";
    } else if (format === "tsv") {
      filename += ".tsv";
      content = exportContactsDelimited(contacts, "\t");
      mime = "text/tab-separated-values;charset=utf-8";
    } else {
      filename += ".csv";
      content = exportContactsDelimited(contacts, ",");
      mime = "text/csv;charset=utf-8";
    }
    downloadFile(filename, content, mime);
    toast(
      "Exported " +
        contacts.length +
        " contact" +
        (contacts.length === 1 ? "" : "s"),
      format.toUpperCase() + " downloaded.",
      "success",
    );
  }
  function loadFile(file) {
    const name = file.name || "contacts";
    const ext = (name.split(".").pop() || "").toLowerCase();
    if (XLSX_EXTS.indexOf(ext) !== -1) {
      const reader = new FileReader();
      reader.onerror = () => toast("Couldn't read that file", "", "error");
      reader.onload = () => {
        let parsed;
        try {
          parsed = parseWorkbookBuffer(reader.result);
        } catch (err) {
          toast(
            "Couldn't parse that spreadsheet",
            (err && err.message) || "Unsupported file.",
            "error",
          );
          return;
        }
        if (!parsed || !parsed.objects.length) {
          toast("No contacts found", "", "warning");
          return;
        }
        applyParsedFile(parsed, name);
      };
      reader.readAsArrayBuffer(file);
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => toast("Couldn't read that file", "", "error");
    reader.onload = () => {
      let parsed;
      try {
        parsed = parseAnyFile(stripBOM(String(reader.result || "")), ext);
      } catch (err) {
        toast(
          "Couldn't parse that file",
          (err && err.message) || "Unsupported file.",
          "error",
        );
        return;
      }
      if (!parsed || !parsed.objects.length) {
        toast("No contacts found", "", "warning");
        return;
      }
      applyParsedFile(parsed, name);
    };
    reader.readAsText(file);
  }
  function applyParsedFile(parsed, name) {
    const doApply = (mappedColMap) => {
      pushHistory();
      state.columns = parsed.columns.slice();
      state.colMap = Object.assign({}, mappedColMap || parsed.colMap);
      state.loadedName = name;
      state.lists = [{ id: "list_1", name: "Main" }];
      state.panes.left.listId = "list_1";
      state.panes.right.listId = null;
      state.panes.left.search = "";
      state.panes.right.search = "";
      state.panes.left.selectedSids.clear();
      state.panes.right.selectedSids.clear();
      state.contacts = parsed.objects.map((values, i) =>
        hydrateContact(
          { values: Object.assign({}, values), listId: "list_1" },
          i,
        ),
      );
      state.editingId = null;
      state.detailsContactSid = null;
      computeAllDuplicates();
      renderAll();
      flushSave();
      toast(
        "Loaded " +
          state.contacts.length +
          " contact" +
          (state.contacts.length === 1 ? "" : "s"),
        "Click ＋ New list above the right pane (or drag rows) to split.",
        "success",
      );
    };
    // No mapping step: the columns are auto-detected and the import goes straight in.
    const importWith = (mappedColMap) => {
      if (state.contacts.length) {
        confirmAction({
          title: "Replace current data?",
          html:
            "You have <b>" +
            state.contacts.length +
            "</b> contact(s). Loading <b>" +
            escapeHtml(name) +
            "</b> will replace them.<br>You can undo with <b>Ctrl+Z</b>.",
          confirmText: "Replace",
          cancelText: "Keep current",
          icon: "question",
        }).then((ok) => {
          if (ok) doApply(mappedColMap);
        });
      } else doApply(mappedColMap);
    };
    importWith(parsed.colMap);
  }
  const MAP_OPTIONS = [
    ["ignore", "Ignore"],
    ["name", "Name"],
    ["surname", "Surname"],
    ["phone", "Phone number"],
    ["phonelabel", "Phone label"],
    ["email", "Email"],
    ["org", "Organization"],
    ["title", "Job title"],
    ["address", "Address"],
    ["note", "Notes"],
    ["other", "Other"],
  ];
  function smartMapColumn(col, current) {
    const n = normalizeHeader(col);
    if (/^e ?mail\s*\d+\s*[-–]?\s*label$/i.test(n)) return "ignore";
    if (
      /^(phone|mobile|cell|telephone|tel|whatsapp|fax|pager)\s*\d+\s*[-–]?\s*label$/i.test(
        n,
      )
    )
      return "phonelabel";
    if (
      /^(phone|mobile|cell|telephone|tel|whatsapp|fax|pager)\s*\d+\s*[-–]?\s*value$/i.test(
        n,
      )
    )
      return "phone";
    if (/^e ?mail\s*\d+\s*[-–]?\s*value$/i.test(n)) return "email";
    if (/^labels?$/i.test(n)) return "ignore";
    const d = current || detectField(col);
    if (d === "first" || d === "middle") return "name";
    if (d === "last") return "surname";
    return d || "other";
  }
  function openColumnMapping(parsed, name, onDone) {
    const old = document.getElementById("columnMappingModal");
    if (old) old.remove();
    const overlay = document.createElement("div");
    overlay.id = "columnMappingModal";
    overlay.className = "mapping-overlay";
    overlay.innerHTML =
      '<div class="mapping-modal" role="dialog" aria-modal="true">' +
      '<div class="mapping-head"><div><div class="mapping-eyebrow">IMPORT · COLUMN MAPPING</div>' +
      "<h3>Check your contact fields</h3><p>" +
      escapeHtml(name) +
      " · " +
      parsed.objects.length +
      " rows</p></div>" +
      '<button type="button" class="mapping-close" data-map-cancel>×</button></div>' +
      '<div class="mapping-summary"><b>' +
      parsed.columns.length +
      "</b> columns detected <span>•</span> Review before import.</div>" +
      '<div class="mapping-table-wrap"><table class="mapping-table"><thead><tr><th>Source column</th><th>Sample</th><th>Use as</th></tr></thead><tbody></tbody></table></div>' +
      '<div class="mapping-foot"><button type="button" class="map-secondary" data-map-auto>Auto-map again</button>' +
      '<span class="mapping-spacer"></span>' +
      '<button type="button" class="map-secondary" data-map-cancel>Cancel</button>' +
      '<button type="button" class="map-primary" data-map-apply>Import contacts</button></div></div>';
    document.body.appendChild(overlay);
    const tbody = overlay.querySelector("tbody");
    const selects = new Map();
    const buildRows = () => {
      tbody.innerHTML = "";
      parsed.columns.forEach((col) => {
        const tr = document.createElement("tr");
        const tdName = document.createElement("td");
        tdName.className = "map-source";
        tdName.textContent = col;
        tdName.title = col;
        const tdSample = document.createElement("td");
        tdSample.className = "map-sample";
        const sample = parsed.objects.find((o) => String(o[col] || "").trim());
        tdSample.textContent = sample ? String(sample[col]).slice(0, 70) : "—";
        tdSample.title = sample ? String(sample[col]) : "";
        const tdSelect = document.createElement("td");
        const select = document.createElement("select");
        select.className = "map-select";
        MAP_OPTIONS.forEach(([value, label]) => {
          const opt = document.createElement("option");
          opt.value = value;
          opt.textContent = label;
          select.appendChild(opt);
        });
        select.value = smartMapColumn(col, parsed.colMap[col]);
        selects.set(col, select);
        tdSelect.appendChild(select);
        tr.append(tdName, tdSample, tdSelect);
        tbody.appendChild(tr);
      });
    };
    buildRows();
    const close = () => overlay.remove();
    overlay
      .querySelectorAll("[data-map-cancel]")
      .forEach((b) => b.addEventListener("click", close));
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    overlay
      .querySelector("[data-map-auto]")
      .addEventListener("click", () => buildRows());
    overlay.querySelector("[data-map-apply]").addEventListener("click", () => {
      const mapped = {};
      parsed.columns.forEach((col) => {
        mapped[col] = selects.get(col).value;
      });
      close();
      onDone(mapped);
    });
  }

  /* ========== 17. DELETE ========== */
  function deleteContacts(sids, sourceSide) {
    const list = (sids || []).filter(Boolean);
    if (!list.length) return;
    pushHistory();
    const kill = new Set(list);
    state.contacts = state.contacts.filter((c) => !kill.has(c._sid));
    if (sourceSide) state.panes[sourceSide].selectedSids.clear();
    rebuildIds();
    computeAllDuplicates();
    renderAll();
    flushSave();
  }
  function deleteSingleContact(sid) {
    const c = state.contacts.find((x) => x._sid === sid);
    if (!c) return;
    const label = c.name || "(no name)";
    confirmAction({
      title: "Delete this contact?",
      html:
        "<b>" +
        escapeHtml(label) +
        "</b> will be removed.<br>You can undo this with <b>Ctrl+Z</b>.",
      confirmText: "Delete",
      fallback: 'Delete "' + label + '"?',
    }).then((ok) => {
      if (ok) deleteContacts([sid]);
    });
  }

  /* ========== 18. EVENTS ========== */
  const closeAllMenus = () => {
    document.querySelectorAll(".export-menu").forEach((m) => {
      m.hidden = true;
    });
  };

  els.fileInput.addEventListener("change", (e) => {
    const f = e.target.files && e.target.files[0];
    e.target.value = "";
    if (f) loadFile(f);
  });

  {
    const side = "left";
    const pe = paneEls(side);



    pe.search.addEventListener("input", () => {
      state.panes[side].search = pe.search.value;
      renderPane(side);
      scheduleSave();
    });

    pe.addBtn.addEventListener("click", () =>
      openContactModal(null, state.panes[side].listId || state.lists[0].id),
    );

    pe.exportBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const open = pe.exportMenu.hidden;
      closeAllMenus();
      pe.exportMenu.hidden = !open;
    });
    pe.exportMenu.addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-export-format]");
      if (!btn) return;
      const contacts = state.panes[side].listId
        ? state.contacts.filter((c) => c.listId === state.panes[side].listId)
        : [];
      exportContacts(contacts, btn.dataset.exportFormat);
      closeAllMenus();
    });

    pe.delSelBtn.addEventListener("click", () => {
      const sids = Array.from(state.panes[side].selectedSids);
      if (!sids.length) return;
      confirmAction({
        title: "Delete selected contacts?",
        html:
          "<b>" +
          sids.length +
          "</b> contact(s) will be removed.<br>You can undo with <b>Ctrl+Z</b>.",
        confirmText: "Delete",
      }).then((ok) => {
        if (ok) deleteContacts(sids, side);
      });
    });
    pe.clearSelBtn.addEventListener("click", () => {
      state.panes[side].selectedSids.clear();
      renderPane(side);
    });

    /* === ROW CLICK — opens DETAILS (unless on a control) === */
    pe.tbody.addEventListener("click", (e) => {
      // 1. Ignore clicks on checkboxes
      if (e.target.closest && e.target.closest("input[type=checkbox]")) return;

      // 2. Action buttons: edit, dedupe, delete
      const btn = e.target.closest
        ? e.target.closest("button[data-act]")
        : null;
      if (btn) {
        const id = btn.dataset.id,
          act = btn.dataset.act;
        if (act === "edit") openContactModal(id);
        else if (act === "delete") deleteSingleContact(id);
        else if (act === "dedupe") dedupeRow(id);
        return;
      }

      // 3. Anywhere else on the row → details modal
      const tr = e.target.closest ? e.target.closest("tr[data-sid]") : null;
      if (!tr) return;
      const sid = tr.dataset.sid;
      if (!sid) return;
      openContactDetails(sid);
    });

  }


  const bindMatch = (el, key) =>
    el &&
    el.addEventListener("change", () => {
      state.match[key] = el.checked;
      computeAllDuplicates();
      renderAll();
      scheduleSave();
    });
  bindMatch(els.matchPhone, "phone");
  bindMatch(els.matchName, "name");
  bindMatch(els.matchEmail, "email");
  if (els.onlyNoNumber)
    els.onlyNoNumber.addEventListener("change", () => {
      state.onlyNoNumber = els.onlyNoNumber.checked;
      renderAll();
      scheduleSave();
    });
  if (els.onlySelfDup)
    els.onlySelfDup.addEventListener("change", () => {
      state.onlySelfDup = els.onlySelfDup.checked;
      renderAll();
      scheduleSave();
    });

  els.modalForm.addEventListener("submit", saveContact);
  els.modalClose.addEventListener("click", closeContactModal);
  els.modalCancel.addEventListener("click", closeContactModal);
  els.contactModal.addEventListener("mousedown", (e) => {
    if (e.target === els.contactModal) closeContactModal();
  });
  els.detailsModalClose.addEventListener("click", closeContactDetails);
  els.detailsModalCloseBtn.addEventListener("click", closeContactDetails);
  els.detailsModal.addEventListener("mousedown", (e) => {
    if (e.target === els.detailsModal) closeContactDetails();
  });
  els.detailsModalEditBtn.addEventListener("click", () => {
    const sid = state.detailsContactSid;
    if (!sid) return;
    closeContactDetails();
    openContactModal(sid);
  });
  els.detailsModalDeleteBtn.addEventListener("click", () => {
    const sid = state.detailsContactSid;
    if (!sid) return;
    closeContactDetails();
    deleteSingleContact(sid);
  });
  if (els.detailsModalDedupeBtn) {
    els.detailsModalDedupeBtn.addEventListener("click", () => {
      const sid = state.detailsContactSid;
      if (!sid) return;
      closeContactDetails();
      dedupeRow(sid);
    });
  }
  els.detailsModalLabelBtn.addEventListener("click", () => {
    const sid = state.detailsContactSid;
    if (!sid) return;
    closeContactDetails();
    autoLabelRow(sid);
  });

  document.addEventListener("click", (e) => {
    if (!e.target.closest(".export-dropdown")) closeAllMenus();
  });

  if (els.saveNowBtn) {
    els.saveNowBtn.addEventListener("click", () => flushSave());
  }
  els.resetMemoryBtn.addEventListener("click", resetMemory);

  document.addEventListener("keydown", (e) => {
    const t = e.target;
    const tag = (t && t.tagName) || "";
    const typing =
      /^(INPUT|TEXTAREA|SELECT)$/.test(tag) || (t && t.isContentEditable);
    const mod = e.ctrlKey || e.metaKey;
    const key = (e.key || "").toLowerCase();
    if (mod && key === "s") {
      e.preventDefault();
      flushSave();
      return;
    }
    if (mod && key === "z" && !e.shiftKey) {
      if (typing) return;
      e.preventDefault();
      undo();
      return;
    }
    if (mod && (key === "y" || (key === "z" && e.shiftKey))) {
      if (typing) return;
      e.preventDefault();
      redo();
      return;
    }
    if (e.key === "Escape") {
      if (!els.contactModal.hidden) {
        closeContactModal();
        return;
      }
      if (!els.detailsModal.hidden) {
        closeContactDetails();
        return;
      }
      closeAllMenus();
    }
  });

  let dragDepth = 0;
  const isFileDrag = (e) => {
    if (!e.dataTransfer) return false;
    const types = e.dataTransfer.types || [];
    return Array.prototype.indexOf.call(types, "Files") !== -1;
  };
  window.addEventListener("dragenter", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    dragDepth++;
    els.dropOverlay.classList.add("show");
  });
  window.addEventListener("dragover", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  });
  window.addEventListener("dragleave", () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) els.dropOverlay.classList.remove("show");
  });
  window.addEventListener("drop", (e) => {
    if (!isFileDrag(e)) return;
    const files = e.dataTransfer.files;
    if (!files || !files.length) return;
    e.preventDefault();
    dragDepth = 0;
    els.dropOverlay.classList.remove("show");
    const f = files[0];
    if (f) loadFile(f);
  });

  window.addEventListener("beforeunload", () => {
    flushSave();
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") flushSave();
  });
  window.addEventListener("pagehide", flushSave);

  /* ========== 19. INIT ========== */
  function init() {
    const had = loadStoredState();
    if (els.matchPhone) els.matchPhone.checked = state.match.phone;
    if (els.matchName) els.matchName.checked = state.match.name;
    if (els.matchEmail) els.matchEmail.checked = state.match.email;
    if (els.onlyNoNumber) els.onlyNoNumber.checked = state.onlyNoNumber;
    if (els.onlySelfDup) els.onlySelfDup.checked = state.onlySelfDup;
    renderAll();

    if (had && (state.contacts.length || state.columns.length)) {
      setSaveStatus(
        "ok",
        "✓ Restored " +
          state.contacts.length +
          " contact" +
          (state.contacts.length === 1 ? "" : "s"),
        true,
      );
    } else {
      setSaveStatus("ok", "✓ Ready (auto-save on)", true);
    }
  }
  init();
})();
