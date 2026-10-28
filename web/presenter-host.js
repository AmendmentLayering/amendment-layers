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
    const url = new URL("presenter.html", window.location.href).href;
    // Add-in windows have no maximize button, so open it filling the screen
    // Word is on; Windows+Shift+arrow then moves it whole to the projector.
    Office.context.ui.displayDialogAsync(url, { height: 100, width: 100, promptBeforeOpen: false }, (res) => {
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
    const first = upTo.items[Math.max(0, at - CONTEXT)];
    const last = from.items[Math.min(from.items.length - 1, CONTEXT)];
    const ox = first.getRange("Start").expandTo(last.getRange("End")).getOoxml();
    const counts = notes ? notesBefore(context, first) : null;
    await context.sync();

    const view = AmdOoxml.presenterView(ox.value, counts ? { notesBefore: counts() } : {});
    return {
      type: "view",
      html: view.html,
      texts: view.texts,
      cursor: Math.min(at, CONTEXT),
      cursorText: here.text,
    };
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
      schedule(50);
    },
  };
})();
