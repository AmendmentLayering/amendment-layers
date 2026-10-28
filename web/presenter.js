/*
 * Presenter view window (opened from the pane with displayDialogAsync).
 * Word sends the passage around the cursor; this page shows it in large type
 * and keeps the current paragraph in the middle of the screen. Nothing here
 * talks to Word directly; the pane does that and sends messages.
 *
 * Message from the pane: { type: "view", html, texts, cursor, cursorText },
 * possibly split into { type: "chunk", id, i, n, data } pieces. With the
 * cursor in a footnote or endnote, html is the main-text paragraph holding
 * that note's number, and `note` is { label, html, texts, cursor,
 * cursorText } for the note, shown in a panel below it.
 */
(() => {
  "use strict";

  const SIZE_KEY = "amendmentLayers.presenterSize";
  const LEGEND_KEY = "amendmentLayers.presenterLegend";
  const MIN = 18;
  const MAX = 96;

  let size = Number(read(SIZE_KEY)) || 34;
  let lastHtml = null;
  const chunks = new Map();

  function read(key) {
    try { return localStorage.getItem(key); } catch (e) { return null; }
  }

  function write(key, value) {
    try { localStorage.setItem(key, String(value)); } catch (e) { /* not saved; fine */ }
  }

  function setSize(px) {
    size = Math.max(MIN, Math.min(MAX, px));
    document.documentElement.style.setProperty("--size", `${size}px`);
    write(SIZE_KEY, size);
    const here = document.querySelector(".here");
    if (here) bringIntoView(here, "auto");
  }

  // In a note, show the main-text paragraph above it too when both fit;
  // otherwise keep the note's current paragraph in the middle.
  function bringIntoView(target, behavior) {
    const ref = target.closest(".note") && document.querySelector(".ref-para");
    if (ref) {
      const top = ref.getBoundingClientRect().top;
      const bottom = target.getBoundingClientRect().bottom;
      const room = window.innerHeight * 0.88 - legendHeight();
      if (bottom - top <= room) {
        window.scrollBy({ top: top - window.innerHeight * 0.05, behavior });
        return;
      }
    }
    target.scrollIntoView({ block: "center", behavior });
  }

  function legendHeight() {
    const legend = document.getElementById("legend");
    return legend && !document.body.classList.contains("no-legend") ? legend.offsetHeight : 0;
  }

  function toggleLegend() {
    document.body.classList.toggle("no-legend");
    write(LEGEND_KEY, document.body.classList.contains("no-legend") ? "off" : "on");
  }

  // The block for the cursor's paragraph: by position, or, if the passage
  // came back with a different paragraph there, by its text.
  function findTarget(doc, view) {
    if (view.cursor < 0) return null;
    const norm = (s) => String(s).replace(/[\s​]+/g, " ").trim();
    const byIndex = doc.querySelector(`[data-i="${view.cursor}"]`);
    if (!view.texts || view.cursorText == null) return byIndex;
    const want = norm(view.cursorText);
    if (norm(view.texts[view.cursor]) === want) return byIndex;
    let best = null;
    view.texts.forEach((t, i) => {
      if (norm(t) === want && (best === null || Math.abs(i - view.cursor) < Math.abs(best - view.cursor))) best = i;
    });
    return best === null ? byIndex : doc.querySelector(`[data-i="${best}"]`);
  }

  function show(view) {
    const doc = document.getElementById("doc");
    const prev = doc.querySelector(".here");
    const prevText = prev ? prev.textContent : null;
    const prevTop = prev ? prev.getBoundingClientRect().top : null;

    const html = view.note
      ? `<div class="ref-para">${view.html}</div>` +
        `<section class="note"><div class="note-label">${escape(view.note.label || "Note")}</div>` +
        `<div class="note-text">${view.note.html}</div></section>`
      : view.html;
    if (html !== lastHtml) {
      doc.innerHTML = html;
      lastHtml = html;
      // Keep the paragraph that was on screen where it was, so a refresh
      // doesn't jump; then glide to the new one.
      if (prevText !== null) {
        const same = [...doc.querySelectorAll(".p")].find((el) => el.textContent === prevText);
        if (same) window.scrollBy(0, same.getBoundingClientRect().top - prevTop);
      }
    }
    for (const el of doc.querySelectorAll(".here")) el.classList.remove("here");
    const target = view.note
      ? findTarget(doc.querySelector(".note-text"), view.note)
      : findTarget(doc, view);
    if (target) {
      target.classList.add("here");
      bringIntoView(target, prev ? "smooth" : "auto");
    }
    document.body.classList.remove("waiting");
  }

  function escape(s) {
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  function onMessage(raw) {
    let msg;
    try { msg = JSON.parse(raw); } catch (e) { return; }
    if (msg.type === "chunk") {
      const parts = chunks.get(msg.id) || [];
      parts[msg.i] = msg.data;
      chunks.set(msg.id, parts);
      if (parts.filter((p) => p !== undefined).length === msg.n) {
        chunks.delete(msg.id);
        onMessage(parts.join(""));
      }
      return;
    }
    if (msg.type === "view") show(msg);
  }

  // Controls appear while the mouse moves, then fade.
  let idle;
  document.addEventListener("mousemove", () => {
    document.body.classList.add("show-controls");
    clearTimeout(idle);
    idle = setTimeout(() => document.body.classList.remove("show-controls"), 2500);
  });
  document.getElementById("controls").addEventListener("click", (e) => {
    const cmd = e.target.closest("button") && e.target.closest("button").dataset.cmd;
    if (cmd === "smaller") setSize(size - 4);
    else if (cmd === "larger") setSize(size + 4);
    else if (cmd === "legend") toggleLegend();
  });
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.altKey || e.metaKey) return;
    if (e.key === "+" || e.key === "=") setSize(size + 4);
    else if (e.key === "-") setSize(size - 4);
    else if (e.key.toLowerCase() === "l") toggleLegend();
  });

  setSize(size);
  if (read(LEGEND_KEY) === "off") document.body.classList.add("no-legend");

  // For previewing this page outside Word.
  window.presenterShow = show;

  Office.onReady(() => {
    if (!Office.context.ui || !Office.context.ui.messageParent) return;
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => onMessage(arg.message));
    Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
  });
})();
