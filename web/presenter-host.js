/*
 * Pane side of Presenter view: opens the presenter window, reads the passage
 * around the cursor from Word, and sends it to the window. It refreshes after
 * every add-in action and, while the window is open, shortly after the
 * selection settles (which also covers ordinary typing).
 */
const Presenter = (() => {
  "use strict";

  const CONTEXT = 12;   // paragraphs shown before and after the cursor's
  const CHUNK = 24000;  // characters per message to the window
  const SETTLE = 350;   // ms after the selection stops moving

  let dialog = null;
  let ready = false;
  let timer = null;
  let running = false;
  let again = false;
  let sendSeq = 0;
  let serial = (fn) => fn();
  let onStatus = () => {};
  let onOpenChange = () => {};

  // ------------------------------------------------------------- the window

  function open() {
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) {
      onStatus("This version of Word can't run Presenter view (it needs Dialog API 1.2).", "error");
      return;
    }
    // Word titles the window with its address, so the address names the
    // document: ...presenter.html?document=Bylaws_Amendments.docx
    const url = new URL("presenter.html", window.location.href);
    url.search = `document=${documentName()}`;
    // Add-in windows have no maximize button, so open it filling the screen
    // Word is on; Windows+Shift+arrow then moves it whole to the projector.
    Office.context.ui.displayDialogAsync(url.href, { height: 100, width: 100, promptBeforeOpen: false }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        onStatus(`Presenter view couldn't open: ${res.error.message}`, "error");
        return;
      }
      dialog = res.value;
      ready = false;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, onWindowMessage);
      dialog.addEventHandler(Office.EventType.DialogEventReceived, closed); // closed by the user
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelection);
      onOpenChange(true);
      onStatus("Presenter view is open. Click its title bar, then press Windows key + Shift + → (or ←) to move it to the projected screen.");
    });
  }

  // The document's file name, with anything but letters, digits, dots and
  // hyphens shown as "_" so the address stays readable (no %20).
  function documentName() {
    let name = "";
    try {
      const where = Office.context.document.url || "";
      name = decodeURIComponent(where.split(/[\\/]/).pop().split(/[?#]/)[0]);
    } catch (e) {
      // An unusual address; fall back below.
    }
    return (name || "Unsaved document").replace(/[^A-Za-z0-9.-]+/g, "_");
  }

  function close() {
    if (dialog) {
      try { dialog.close(); } catch (e) { /* already gone */ }
    }
    closed();
  }

  function closed() {
    if (!dialog && !ready) return;
    dialog = null;
    ready = false;
    clearTimeout(timer);
    Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: onSelection });
    onOpenChange(false);
  }

  function onWindowMessage(arg) {
    let msg;
    try { msg = JSON.parse(arg.message); } catch (e) { return; }
    if (msg.type === "ready") {
      ready = true;
      refresh();
    }
  }

  function onSelection() {
    schedule(SETTLE);
  }

  // ------------------------------------------------------------- refreshing

  function schedule(delay) {
    if (!dialog) return;
    clearTimeout(timer);
    timer = setTimeout(refresh, delay);
  }

  // One refresh at a time; a request during one runs once it finishes.
  function refresh() {
    if (!dialog || !ready) return;
    if (running) {
      again = true;
      return;
    }
    running = true;
    serial(capture)
      .then(send)
      .catch((err) => console.warn("Presenter view refresh failed", err))
      .then(() => {
        running = false;
        if (again) {
          again = false;
          refresh();
        }
      });
  }

  // Footnotes and endnotes need WordApi 1.5. Without it the window shows the
  // main text only, without note numbers.
  const notesSupported = () => Office.context.requirements.isSetSupported("WordApi", "1.5");
  const NOTE_BODY = new Set(["Footnote", "Endnote", "NoteItem"]);
  const INSIDE = new Set(["Inside", "InsideStart", "InsideEnd", "Equal"]);

  // Footnotes and endnotes earlier in the document than `range`, so the
  // window can number the notes in the passage as Word does.
  function notesBefore(context, range) {
    const before = context.document.body.getRange("Start").expandTo(range.getRange("Start"));
    const fn = before.footnotes;
    const en = before.endnotes;
    fn.load("items/type");
    en.load("items/type");
    return () => ({ footnote: fn.items.length, endnote: en.items.length });
  }

  // The passage: CONTEXT paragraphs either side of the cursor's paragraph.
  // With the cursor in a footnote or endnote: the main-text paragraph with
  // that note's number, and below it the note.
  function capture() {
    return Word.run(async (context) => {
      const notes = notesSupported();
      const sel = context.document.getSelection();
      const here = sel.paragraphs.getFirst();
      here.load("text");
      let where = null;
      if (notes) {
        where = sel.parentBody;
        where.load("type");
      }
      await context.sync();
      if (notes && NOTE_BODY.has(where.type)) {
        let view = null;
        try {
          view = await captureNote(context, sel, here);
        } catch (err) {
          console.warn("Presenter view couldn't read the note", err);
        }
        if (view) return view;
      }
      return captureMain(context, here, notes);
    });
  }

  async function captureMain(context, here, notes) {
    const body = context.document.body;
    const upTo = body.getRange("Start").expandTo(here.getRange("End")).paragraphs; // ends with `here`
    const from = here.getRange("Start").expandTo(body.getRange("End")).paragraphs; // starts with `here`
    upTo.load("items/tableNestingLevel");
    from.load("items/tableNestingLevel");
    await context.sync();

    const at = upTo.items.length - 1;
    const firstAt = Math.max(0, at - CONTEXT);
    const first = upTo.items[firstAt];
    const lastAt = Math.min(from.items.length - 1, CONTEXT);
    const last = from.items[lastAt];
    const shown = upTo.items.slice(firstAt).concat(from.items.slice(1, lastAt + 1));
    const ox = first.getRange("Start").expandTo(last.getRange("End")).getOoxml();
    const counts = notes ? notesBefore(context, first) : null;
    const listed = listsSupported() ? shown.map((p) => {
      const li = p.listItemOrNullObject;
      li.load("listString");
      return li;
    }) : null;
    await context.sync();

    const total = upTo.items.length + from.items.length - 1;
    // Up to and including the last paragraph mark: a paragraph's list
    // setting is on its mark, so without it the last paragraph has no number.
    const whole = () => upTo.items[0].getRange("Start").expandTo(from.items[from.items.length - 1].getRange("Whole"));
    const numbers = listed ? await listNumbers(context, listed, firstAt, total, whole) : null;
    const opts = counts ? { notesBefore: counts() } : {};
    if (numbers) opts.numbers = numbers;
    const view = AmdOoxml.presenterView(ox.value, opts);
    return {
      type: "view",
      html: view.html,
      texts: view.texts,
      cursor: Math.min(at, CONTEXT),
      cursorText: here.text,
    };
  }

  // ------------------------------------------------------------ list numbers

  // Word's automatic list numbers need WordApi 1.3 (Paragraph.listItem).
  const listsSupported = () => Office.context.requirements.isSetSupported("WordApi", "1.3");
  const RECHECK = 10000; // ms before a mismatch with Word's numbers is worked out again

  // The whole document's numbering, worked out from its OOXML: kept between
  // refreshes, as reading the whole document is slow, and worked out again
  // after an add-in action, when the paragraph count changes, or when it no
  // longer matches the numbers Word shows.
  let numbering = null; // { total, list, at }
  let numberingDirty = true;

  // Word's number for each shown paragraph (`listed`, from the paragraph
  // at body index `firstAt`), with the before/after numbers when the
  // add-in's arithmetic agrees with Word's for every shown paragraph.
  // `whole()` is the range from the first paragraph to the last, read the
  // same way as the shown passage, so its paragraphs line up with Word's.
  async function listNumbers(context, listed, firstAt, total, whole) {
    const words = listed.map((li) => (li.isNullObject ? null : li.listString));
    // A package may end with an extra empty paragraph; the check of every
    // shown number against Word's catches anything else out of line.
    const lined = () => numbering && numbering.list.length >= total;
    const agrees = () => lined() && words.every((w, j) => {
      const n = numbering.list[firstAt + j];
      if (w == null || !n) return w == null && !n;
      return n.word === "•" || n.word === w;
    });
    const stale = !numbering || numberingDirty || numbering.total !== total ||
      (!agrees() && Date.now() - numbering.at > RECHECK);
    let error = null;
    if (stale && words.some((w) => w != null)) {
      try {
        const ox = whole().getOoxml();
        await context.sync();
        numbering = { total, list: AmdOoxml.listNumbers(ox.value), at: Date.now() };
        numberingDirty = false;
      } catch (err) {
        console.warn("Presenter view couldn't work out the list numbers", err);
        error = String((err && err.message) || err);
        numbering = null;
      }
      noteNumbers(words, firstAt, total, error, agrees());
    }
    const aligned = lined();
    const exact = aligned && agrees();
    return words.map((w, j) => {
      if (w == null) return null;
      const n = aligned ? numbering.list[firstAt + j] : null;
      const text = n && n.word === "•" ? "•" : w.replace(/[-]/g, "•");
      if (!n) return { text };
      const out = { text, status: n.status, level: n.level };
      if (exact) Object.assign(out, { orig: n.orig, prop: n.prop, cause: n.cause });
      return out;
    });
  }

  // A diagnostic note in the document's add-in settings (read with
  // test\dump-docx.ps1): how the last working-out of the numbers compared
  // with Word's.
  function noteNumbers(words, firstAt, total, error, agreed) {
    try {
      const list = numbering ? numbering.list : [];
      const shown = words.map((w, j) => {
        const n = list[firstAt + j];
        return w == null && !n ? null : `${w}|${n ? `${n.word}>${n.prop}${n.status ? ` ${n.status}${n.level}` : ""}` : "-"}`;
      }).filter(Boolean);
      const value = { total, worked: numbering ? list.length : null, agreed, error, shown };
      const settings = Office.context.document.settings;
      settings.set("amendmentLayersLastNumbers", { at: new Date().toISOString(), value });
      settings.saveAsync(() => {});
    } catch (e) {
      // Diagnostics only.
    }
  }

  // Null if the note holding the cursor can't be found (the main text is
  // shown instead).
  async function captureNote(context, sel, here) {
    const body = context.document.body;
    const lists = [["footnote", body.footnotes], ["endnote", body.endnotes]];
    for (const [, list] of lists) list.load("items/type");
    await context.sync();

    const checks = [];
    for (const [kind, list] of lists) {
      list.items.forEach((note, i) => checks.push({ kind, num: i + 1, note, rel: sel.compareLocationWith(note.body.getRange("Whole")) }));
    }
    await context.sync();
    const hit = checks.find((c) => INSIDE.has(c.rel.value));
    if (!hit) return null;

    const refPara = hit.note.reference.paragraphs.getFirst();
    const refOx = refPara.getRange("Whole").getOoxml();
    const counts = notesBefore(context, refPara);
    const noteOx = hit.note.body.getOoxml();
    const upTo = hit.note.body.getRange("Start").expandTo(here.getRange("End")).paragraphs;
    upTo.load("items/tableNestingLevel");
    await context.sync();

    const current = { kind: hit.kind, num: hit.num };
    const main = AmdOoxml.presenterView(refOx.value, { notesBefore: counts(), current });
    const note = AmdOoxml.presenterView(noteOx.value, { notesBefore: counts(), ownNote: current, fmt: main.fmt });
    return {
      type: "view",
      html: main.html,
      texts: main.texts,
      cursor: -1, // no main-text paragraph is the one being worked on
      cursorText: null,
      note: {
        label: main.label,
        html: note.html,
        texts: note.texts,
        cursor: upTo.items.length - 1,
        cursorText: here.text,
      },
    };
  }

  function send(view) {
    if (!dialog) return;
    const text = JSON.stringify(view);
    if (text.length <= CHUNK) {
      dialog.messageChild(text);
      return;
    }
    const id = ++sendSeq;
    const n = Math.ceil(text.length / CHUNK);
    for (let i = 0; i < n; i++) {
      dialog.messageChild(JSON.stringify({ type: "chunk", id, i, n, data: text.slice(i * CHUNK, (i + 1) * CHUNK) }));
    }
  }

  return {
    init(options) {
      serial = options.serial || serial;
      onStatus = options.onStatus || onStatus;
      onOpenChange = options.onOpenChange || onOpenChange;
    },
    get isOpen() {
      return !!dialog;
    },
    toggle() {
      if (dialog) close();
      else open();
    },
    // Called after every add-in action.
    changed() {
      numberingDirty = true;
      schedule(50);
    },
  };
})();
