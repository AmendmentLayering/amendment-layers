/*
 * Pure OOXML transforms for amendment-layer formatting.
 *
 * Nothing in this file touches Office.js, so the same code runs inside the
 * add-in and in test/test.html. Every public function takes the flat-OPC
 * package string returned by Word's Range.getOoxml() and returns a package
 * string ready for Range.insertOoxml().
 *
 * Run classification (matches three_layer_amendment_mockup.docx):
 *   Level 1  insertion: single underline        deletion: single strike
 *   Level 2  insertion: blue, double underline  deletion: blue, double strike
 *   Level 3  insertion: purple, dotted underline deletion: purple, strike
 * Levels 2 and 3 carry pale shading; Level 3 is also italic.
 *
 * Text the add-in generates itself (the struck copy of a lower-level insertion,
 * and the plain space separating marked segments) carries <w:noProof/> as an
 * invisible marker, so fold-up can find exactly what it generated.
 */
(function (root) {
  "use strict";

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
  const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const XML_NS = "http://www.w3.org/XML/1998/namespace";

  const LEVEL_COLOR = { 1: "000000", 2: "1F5FBF", 3: "8B1E8B" };
  const LEVEL_SHADE = { 2: "DCE9FA", 3: "F5DDF5" };
  const UNDERLINE = { 1: "single", 2: "double", 3: "dotted" };

  // Run properties this add-in owns. Anything else in w:rPr (fonts, size,
  // bold, ...) is left untouched.
  const MANAGED = new Set(["i", "iCs", "strike", "dstrike", "noProof", "color", "u", "shd"]);

  // Schema order of w:rPr children; Word rejects out-of-order properties.
  const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
    "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
    "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath", "rPrChange"];

  // Package parts not needed to insert or restore a span of text.
  const DROP_PARTS = [/^\/word\/styles(WithEffects)?\.xml$/, /^\/word\/theme\//,
    /^\/word\/fontTable\.xml$/, /^\/word\/webSettings\.xml$/, /^\/word\/settings\.xml$/,
    /^\/docProps\//, /^\/word\/glossary\//, /^\/customXml\//];

  // ---------------------------------------------------------------- DOM helpers

  function parse(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("Could not read the document's OOXML.");
    return doc;
  }

  function serialize(doc) {
    return new XMLSerializer().serializeToString(doc);
  }

  function wEl(doc, name, attrs) {
    const el = doc.createElementNS(W, "w:" + name);
    for (const k of Object.keys(attrs || {})) el.setAttributeNS(W, "w:" + k, attrs[k]);
    return el;
  }

  function kids(el, name) {
    const out = [];
    if (!el) return out;
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && c.namespaceURI === W && (!name || c.localName === name)) out.push(c);
    }
    return out;
  }

  function kid(el, name) {
    return kids(el, name)[0] || null;
  }

  function wVal(el) {
    return el ? el.getAttributeNS(W, "val") || null : null;
  }

  function isOn(el) {
    if (!el) return false;
    const v = wVal(el);
    return !(v === "0" || v === "false" || v === "off");
  }

  function insertAfter(ref, el) {
    ref.parentNode.insertBefore(el, ref.nextSibling);
    return el;
  }

  function documentBody(doc) {
    for (const part of Array.from(doc.getElementsByTagNameNS(PKG, "part"))) {
      if (part.getAttributeNS(PKG, "name") === "/word/document.xml") {
        return part.getElementsByTagNameNS(W, "body")[0];
      }
    }
    return doc.getElementsByTagNameNS(W, "body")[0];
  }

  function sdtTag(sdt) {
    return wVal(kid(kid(sdt, "sdtPr"), "tag"));
  }

  // ------------------------------------------------------------ run semantics

  function classify(r) {
    const rPr = kid(r, "rPr");
    const color = (wVal(kid(rPr, "color")) || "").toUpperCase();
    const u = kid(rPr, "u");
    const underlined = !!u && (wVal(u) || "single") !== "none";
    const struck = isOn(kid(rPr, "strike")) || isOn(kid(rPr, "dstrike"));
    const level = color === LEVEL_COLOR[2] ? 2 : color === LEVEL_COLOR[3] ? 3 : 1;
    const kind = struck ? "del" : underlined ? "ins" : "plain";
    return { kind, level: kind === "plain" ? 0 : level, gen: isOn(kid(rPr, "noProof")) };
  }

  function runText(r) {
    let s = "";
    for (const c of kids(r)) {
      if (c.localName === "t") s += c.textContent;
      else if (c.localName === "tab") s += "\t";
      else if (c.localName === "br" || c.localName === "cr") s += "\n";
      else if (c.localName === "noBreakHyphen") s += "-";
    }
    return s;
  }

  function hasVisibleContent(r) {
    return kids(r).some((c) => c.localName !== "rPr" && (c.localName !== "t" || c.textContent.length));
  }

  function isSep(r) {
    const c = classify(r);
    return c.kind === "plain" && c.gen && /^\s+$/.test(runText(r));
  }

  function isCopyAt(r, level) {
    const c = classify(r);
    return c.kind === "del" && c.gen && c.level === level;
  }

  function levelProps(kind, level) {
    if (kind === "plain" || !level) return [];
    const p = [];
    if (level === 3) p.push(["i"], ["iCs"]);
    if (kind === "del") p.push([level === 2 ? "dstrike" : "strike"]);
    p.push(["color", { val: LEVEL_COLOR[level] }]);
    if (kind === "ins") p.push(["u", { val: UNDERLINE[level] }]);
    if (LEVEL_SHADE[level]) p.push(["shd", { val: "clear", color: "auto", fill: LEVEL_SHADE[level] }]);
    return p;
  }

  function insertOrdered(rPr, el) {
    const idx = RPR_ORDER.indexOf(el.localName);
    for (let c = rPr.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && RPR_ORDER.indexOf(c.localName) > idx) {
        rPr.insertBefore(el, c);
        return;
      }
    }
    rPr.appendChild(el);
  }

  function restyle(r, kind, level, gen) {
    const doc = r.ownerDocument;
    let rPr = kid(r, "rPr");
    if (!rPr) {
      rPr = wEl(doc, "rPr");
      r.insertBefore(rPr, r.firstChild);
    }
    for (const c of kids(rPr)) if (MANAGED.has(c.localName)) rPr.removeChild(c);
    const props = levelProps(kind, level);
    if (gen) props.push(["noProof"]);
    for (const [name, attrs] of props) insertOrdered(rPr, wEl(doc, name, attrs));
  }

  function setRunText(r, text) {
    for (const c of kids(r)) if (c.localName !== "rPr") r.removeChild(c);
    const t = wEl(r.ownerDocument, "t");
    t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.textContent = text;
    r.appendChild(t);
  }

  function makeSep(like) {
    const sep = like.cloneNode(true);
    setRunText(sep, " ");
    restyle(sep, "plain", 0, true);
    return sep;
  }

  // Splits a plain-text run so its first `at` characters stay in `r`; returns
  // the new run holding the rest, or null if the run has non-text content.
  function splitRun(r, at) {
    const content = kids(r).filter((c) => c.localName !== "rPr");
    if (content.length !== 1 || content[0].localName !== "t") return null;
    const text = content[0].textContent;
    const tail = r.cloneNode(true);
    setRunText(r, text.slice(0, at));
    setRunText(tail, text.slice(at));
    return insertAfter(r, tail);
  }

  // Paragraphs in the body, skipping any inside a content control whose tag is
  // in `skip` (another pending amendment's span).
  function paragraphs(doc, skip, found) {
    const out = [];
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === W) {
          if (c.localName === "p") { out.push(c); continue; }
          if (c.localName === "sectPr") continue;
          if (c.localName === "sdt" && skip.has(sdtTag(c))) { found.add(sdtTag(c)); continue; }
        }
        walk(c);
      }
    })(documentBody(doc));
    return out;
  }

  function collectRuns(node, skip, found, out) {
    out = out || [];
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if (c.namespaceURI === W) {
        const n = c.localName;
        if (n === "r") { out.push(c); continue; }
        if (n === "pPr" || n === "rPr" || n === "sdtPr" || n === "del" || n === "moveFrom") continue;
        if (n === "sdt" && skip.has(sdtTag(c))) { found.add(sdtTag(c)); continue; }
      }
      collectRuns(c, skip, found, out);
    }
    return out;
  }

  // Removes `drop` runs from one paragraph's run list. With `tidy`, also
  // removes a doubled space left where the text used to be.
  function removeRuns(runs, drop, tidy) {
    const sites = [];
    for (let i = 0; i < runs.length; i++) {
      if (!drop.has(runs[i])) continue;
      let j = i;
      while (j + 1 < runs.length && drop.has(runs[j + 1])) j++;
      sites.push([runs[i - 1] || null, runs[j + 1] || null]);
      i = j;
    }
    for (const r of drop) if (r.parentNode) r.parentNode.removeChild(r);
    if (!tidy) return;
    for (const [prev, next] of sites) {
      const pt = prev ? runText(prev) : null;
      const nt = next ? runText(next) : null;
      if (prev && next && /\s$/.test(pt) && /^[\s,.;:)\]]/.test(nt)) trimRun(prev, "end");
      else if (!prev && next && /^\s/.test(nt)) trimRun(next, "start");
    }
  }

  function trimRun(r, side) {
    const ts = kids(r, "t");
    const t = side === "end" ? ts[ts.length - 1] : ts[0];
    if (!t) return;
    t.textContent = side === "end" ? t.textContent.replace(/\s$/, "") : t.textContent.replace(/^\s/, "");
    if (!hasVisibleContent(r) && r.parentNode) r.parentNode.removeChild(r);
  }

  // Generated separators survive only between two pieces of text in the same
  // paragraph; any left at an edge of the span, or doubled up, are removed.
  function cleanupSeps(doc, skip) {
    const found = new Set();
    const flat = [];
    for (const p of paragraphs(doc, skip, found)) {
      for (const r of collectRuns(p, skip, found)) flat.push({ r, p });
    }
    const keep = flat.filter((x, i) => {
      if (!isSep(x.r)) return true;
      const prev = flat[i - 1];
      const next = flat[i + 1];
      const ok = prev && next && prev.p === x.p && next.p === x.p && !isSep(next.r);
      if (!ok) x.r.parentNode.removeChild(x.r);
      return ok;
    });
    return keep.length;
  }

  // ------------------------------------------------------------- transforms

  // A selection with spaces on both sides of the words: if both were struck,
  // adopting would join the neighbouring words ("the word is" -> "theis").
  // Keep the leading space plain; returns that run, or null.
  function leadingSpaceToKeep(doc, skip, found) {
    const paras = paragraphs(doc, skip, found);
    if (!paras.length) return null;
    const first = collectRuns(paras[0], skip, found).find((r) => runText(r).length);
    const last = collectRuns(paras[paras.length - 1], skip, found).reverse().find((r) => runText(r).length);
    if (!first || !last || first === last && /^\s+$/.test(runText(first))) return null;
    if (!/^\s/.test(runText(first)) || !/\s$/.test(runText(last)) || classify(first).kind !== "plain") return null;
    const lead = /^\s+/.exec(runText(first))[0].length;
    if (lead < runText(first).length && !splitRun(first, lead)) return null;
    return first;
  }

  function markDeleteDoc(doc, level, skip, found) {
    let changed = 0;
    const keepPlain = leadingSpaceToKeep(doc, skip, found);
    for (const p of paragraphs(doc, skip, found)) {
      const runs = collectRuns(p, skip, found);
      for (let i = 0; i < runs.length;) {
        const r = runs[i];
        const c = classify(r);
        if (c.kind === "ins" && c.level < level && /^\s*$/.test(runText(r))) {
          // A space caught at the edge of the selection: nothing to strike.
          i++;
          continue;
        }
        if (c.kind === "ins" && c.level < level) {
          // A lower level's insertion stays visible; a struck copy of it at
          // this level follows, set off by a generated space.
          const group = [r];
          let j = i + 1;
          while (j < runs.length && runs[j].parentNode === r.parentNode) {
            const cj = classify(runs[j]);
            if (cj.kind !== "ins" || cj.level !== c.level) break;
            group.push(runs[j++]);
          }
          let prev = insertAfter(group[group.length - 1], makeSep(group[group.length - 1]));
          for (const g of group) {
            const copy = g.cloneNode(true);
            restyle(copy, "del", level, true);
            prev = insertAfter(prev, copy);
          }
          changed++;
          i = j;
          continue;
        }
        if (c.kind === "ins" && c.level === level) {
          // Striking this amendment's own insertion just takes it back out.
          r.parentNode.removeChild(r);
          changed++;
        } else if (c.kind === "plain" && r !== keepPlain) {
          restyle(r, "del", level, false);
          changed++;
        }
        i++;
      }
    }
    return changed;
  }

  function restyleAllDoc(doc, kind, level, gen, skip, found) {
    let changed = 0;
    for (const p of paragraphs(doc, skip, found)) {
      for (const r of collectRuns(p, skip, found)) {
        restyle(r, kind, level, gen);
        changed++;
      }
    }
    return changed;
  }

  // Adopted: the amendment at `level` becomes part of the text at level-1.
  function foldDoc(doc, level, skip, found, tidy, problems) {
    const target = level - 1;
    let changed = 0;
    for (const p of paragraphs(doc, skip, found)) {
      const runs = collectRuns(p, skip, found);
      const drop = new Set();

      // 1. Struck copies: the lower-level insertion they strike is either
      //    removed (it belonged to the level being folded into) or the copy
      //    moves down one level.
      for (let i = 0; i < runs.length; i++) {
        if (!isCopyAt(runs[i], level)) continue;
        let j = i;
        let copyText = "";
        while (j < runs.length && isCopyAt(runs[j], level)) copyText += runText(runs[j++]);
        const copies = runs.slice(i, j);

        let k = i - 1;
        const seps = [];
        while (k >= 0 && isSep(runs[k])) seps.push(runs[k--]);
        // Spaces at either end don't count when matching a copy to its
        // original: a selection may or may not have caught the space after a word.
        const want = copyText.trim();
        if (!want) {
          for (const r of seps.concat(copies)) drop.add(r);
          i = j - 1;
          continue;
        }
        const orig = [];
        let acc = "";
        let origLevel = null;
        while (k >= 0 && acc.trim().length < want.length) {
          const ck = classify(runs[k]);
          if (ck.kind !== "ins" || (origLevel !== null && ck.level !== origLevel)) break;
          origLevel = ck.level;
          orig.unshift(runs[k]);
          acc = runText(runs[k]) + acc;
          k--;
        }
        // The copy may strike only the end of a longer insertion: split the
        // first original run so just the struck part is affected.
        const at = acc.lastIndexOf(want);
        if (orig.length && at > 0 && /\S/.test(acc.slice(0, at)) && at < runText(orig[0]).length) {
          const tail = splitRun(orig[0], at);
          if (tail) {
            runs.splice(runs.indexOf(orig[0]) + 1, 0, tail);
            orig[0] = tail;
            acc = acc.slice(at);
            j++;
          }
        }
        if (origLevel === null || acc.trim() !== want) {
          problems.push(`Could not find the text that the struck "${copyText}" refers to.`);
        } else if (origLevel === target) {
          for (const r of orig.concat(seps, copies)) drop.add(r);
        } else {
          for (const r of copies) restyle(r, "del", target, true);
        }
        changed++;
        i = j - 1;
      }

      // 2. In-place deletions and insertions at this level move down a level;
      //    at Level 1 they become final (deleted text goes, inserted text
      //    becomes plain).
      for (const r of runs) {
        if (drop.has(r)) continue;
        const c = classify(r);
        if (c.level !== level || (c.kind === "del" && c.gen)) continue;
        if (c.kind === "del") {
          if (target >= 1) restyle(r, "del", target, false);
          else drop.add(r);
        } else {
          restyle(r, target >= 1 ? "ins" : "plain", target, false);
        }
        changed++;
      }
      removeRuns(runs, drop, tidy);
    }
    cleanupSeps(doc, skip);
    return changed;
  }

  // Failed, for Level 1 markup the add-in did not create (the document as
  // noticed): insertions go, deletions become plain text again.
  function failUntrackedDoc(doc, skip, found) {
    let changed = 0;
    for (const p of paragraphs(doc, skip, found)) {
      const runs = collectRuns(p, skip, found);
      const drop = new Set();
      for (const r of runs) {
        const c = classify(r);
        if (c.level !== 1) continue;
        if (c.kind === "ins") drop.add(r);
        else restyle(r, "plain", 0, false);
        changed++;
      }
      removeRuns(runs, drop, true);
    }
    cleanupSeps(doc, skip);
    return changed;
  }

  // Runs of one paragraph in order, each with the tag of the pending span it
  // sits in (if any). Unlike collectRuns, this descends into every content
  // control, so character offsets match the paragraph's text in Word.
  function runsWithSpans(p, pending) {
    const out = [];
    (function walk(node, span) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === W) {
          const n = c.localName;
          if (n === "r") { out.push({ r: c, span }); continue; }
          if (n === "pPr" || n === "rPr" || n === "sdtPr" || n === "del" || n === "moveFrom") continue;
          if (n === "sdt") {
            const tag = sdtTag(c);
            walk(c, pending.has(tag) ? tag : span);
            continue;
          }
        }
        walk(c, span);
      }
    })(p, null);
    return out;
  }

  // Untracked Level 1 at the cursor. The change at character `offset` is the
  // stretch of marked text around it: struck and inserted words, with only
  // spaces and punctuation (such as a comma left unstruck) between them.
  // Plain words or numbers end it, so a separate change elsewhere
  // in the paragraph is left alone. If a pending amendment's span is part of
  // the stretch, nothing changes and info.blocked names that span.
  function resolveUntrackedAtDoc(doc, offset, adopted, pending, info) {
    const p = paragraphs(doc, new Set(), new Set())[0];
    if (!p) return 0;
    const items = runsWithSpans(p, pending);
    let pos = 0;
    for (const it of items) {
      it.start = pos;
      pos += runText(it.r).length;
      it.end = pos;
    }
    const member = (it) => !!it.span || isSep(it.r) || classify(it.r).kind !== "plain";
    const blank = (it) => !it.span && classify(it.r).kind === "plain" && /^[\s \p{P}]+$/u.test(runText(it.r));

    const idx = items.findIndex((it) => member(it) && it.end > it.start && it.start <= offset && offset <= it.end);
    if (idx < 0) return 0;
    let a = idx;
    let b = idx;
    while (a > 0 && (member(items[a - 1]) || blank(items[a - 1]))) a--;
    while (b < items.length - 1 && (member(items[b + 1]) || blank(items[b + 1]))) b++;
    while (a < idx && blank(items[a])) a++;
    while (b > idx && blank(items[b])) b--;
    const unit = items.slice(a, b + 1);

    const blocking = unit.find((it) => it.span);
    if (blocking) {
      info.blocked = blocking.span;
      return 0;
    }
    const drop = new Set();
    let changed = 0;
    for (const it of unit) {
      const c = classify(it.r);
      if (c.level !== 1 || c.gen) continue;
      if (c.kind === "ins") {
        if (adopted) restyle(it.r, "plain", 0, false);
        else drop.add(it.r);
      } else if (adopted) {
        drop.add(it.r);
      } else {
        restyle(it.r, "plain", 0, false);
      }
      changed++;
    }
    removeRuns(items.map((it) => it.r), drop, true);
    cleanupSeps(doc, pending);
    return changed;
  }

  // ------------------------------------------------------------ presenter view

  const HTML_ESC = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" };
  const esc = (s) => s.replace(/[&<>"]/g, (ch) => HTML_ESC[ch]);

  // Runs of a paragraph for display: inside content controls, hyperlinks and
  // tracked insertions; not tracked deletions.
  function displayRuns(p) {
    const out = [];
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === W) {
          const n = c.localName;
          if (n === "r") { out.push(c); continue; }
          if (n === "pPr" || n === "rPr" || n === "sdtPr" || n === "del" || n === "moveFrom") continue;
        }
        walk(c);
      }
    })(p);
    return out;
  }

  // Note numbers as Word shows them. Word's defaults: footnotes 1, 2, 3;
  // endnotes i, ii, iii.
  const NOTE_DEFAULT_FMT = { footnote: "decimal", endnote: "lowerRoman" };
  const CHICAGO = ["*", "†", "‡", "§"];

  function roman(n) {
    const table = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
      [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
    let s = "";
    for (const [v, r] of table) while (n >= v) { s += r; n -= v; }
    return s;
  }

  function formatNoteNumber(n, fmt) {
    switch (fmt) {
      case "lowerRoman": return roman(n);
      case "upperRoman": return roman(n).toUpperCase();
      case "lowerLetter": return String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
      case "upperLetter": return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
      case "chicago": return CHICAGO[(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1);
      default: return String(n);
    }
  }

  // The document's note number formats: the section's setting, else the
  // document's, else Word's defaults.
  function noteFormats(doc) {
    const fmt = { ...NOTE_DEFAULT_FMT };
    const from = (el) => {
      if (!el) return;
      for (const [kind, name] of [["footnote", "footnotePr"], ["endnote", "endnotePr"]]) {
        const v = wVal(kid(kid(el, name), "numFmt"));
        if (v) fmt[kind] = v;
      }
    };
    for (const part of Array.from(doc.getElementsByTagNameNS(PKG, "part"))) {
      if (part.getAttributeNS(PKG, "name") === "/word/settings.xml") {
        from(part.getElementsByTagNameNS(W, "settings")[0]);
      }
    }
    const body = documentBody(doc);
    from(body && body.getElementsByTagNameNS(W, "sectPr")[0]);
    return fmt;
  }

  // A note reference in the text, or the note's own number at its start.
  function noteMarkHtml(c, state) {
    const kind = /^footnote/.test(c.localName) ? "footnote" : "endnote";
    if (!state.notes) return "";
    let n;
    if (/Ref$/.test(c.localName)) {
      if (!state.ownNote) return "";
      n = state.ownNote.num;
    } else {
      // A custom mark is ordinary text in the next run and takes no number.
      const custom = c.getAttributeNS(W, "customMarkFollows");
      if (custom && !["0", "false", "off"].includes(custom)) return "";
      n = ++state.notes[kind];
    }
    const cur = state.current && state.current.kind === kind && state.current.num === n;
    return `<sup class="nref${cur ? " cur" : ""}">${esc(formatNoteNumber(n, state.fmt[kind]))}</sup>`;
  }

  function runHtml(r, state) {
    const rPr = kid(r, "rPr");
    if (wVal(kid(rPr, "rStyle")) === "PlaceholderText" || isOn(kid(rPr, "vanish"))) return { html: "", text: "" };
    let html = "";
    let text = "";
    for (const c of kids(r)) {
      if (/^(footnote|endnote)Ref(erence)?$/.test(c.localName)) {
        const mark = noteMarkHtml(c, state);
        if (mark) return { html: mark, text: "" };
      } else if (c.localName === "t") {
        const t = c.textContent.replace(/​/g, "");
        html += esc(t);
        text += t;
      } else if (c.localName === "tab") {
        html += " ";
        text += "\t";
      } else if (c.localName === "br" || c.localName === "cr") {
        html += "<br>";
        text += "\n";
      } else if (c.localName === "noBreakHyphen") {
        html += "-";
        text += "-";
      }
    }
    if (!html) return { html: "", text: "" };
    const cls = classify(r);
    const classes = [];
    if (cls.kind !== "plain") classes.push(`a-${cls.kind}${cls.level}`);
    if (isOn(kid(rPr, "b"))) classes.push("b");
    if (isOn(kid(rPr, "i")) && cls.level !== 3) classes.push("i");
    const va = wVal(kid(rPr, "vertAlign"));
    if (va === "superscript") html = `<sup>${html}</sup>`;
    else if (va === "subscript") html = `<sub>${html}</sub>`;
    return { html: classes.length ? `<span class="${classes.join(" ")}">${html}</span>` : html, text };
  }

  // Word's paragraph order, including paragraphs inside tables, so index i
  // here is the i-th paragraph Word reports for the same range.
  function presenterBlocks(node, state) {
    let html = "";
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1 || c.namespaceURI !== W) continue;
      if (c.localName === "p") {
        html += paragraphHtml(c, state);
      } else if (c.localName === "tbl") {
        html += "<table class=\"t\">";
        for (const tr of kids(c, "tr")) {
          html += "<tr>";
          for (const tc of kids(tr, "tc")) html += `<td>${presenterBlocks(tc, state)}</td>`;
          html += "</tr>";
        }
        html += "</table>";
      } else if (c.localName === "sdt") {
        const content = kid(c, "sdtContent");
        if (content) html += presenterBlocks(content, state);
      }
    }
    return html;
  }

  function paragraphHtml(p, state) {
    const pPr = kid(p, "pPr");
    const styleId = wVal(kid(pPr, "pStyle")) || "";
    const heading = /^Heading([1-6])$/.exec(styleId);
    const classes = ["p"];
    if (heading) classes.push(`h${heading[1]}`);
    else if (styleId === "Title") classes.push("h1");
    const styles = [];
    const jc = wVal(kid(pPr, "jc"));
    if (jc === "center" || jc === "right") styles.push(`text-align:${jc}`);
    const ind = kid(pPr, "ind");
    const left = ind ? Number(ind.getAttributeNS(W, "left") || ind.getAttributeNS(W, "start") || 0) : 0;
    if (left > 0) styles.push(`margin-left:${(left / 240).toFixed(2)}em`);

    let html = "";
    let text = "";
    for (const r of displayRuns(p)) {
      const piece = runHtml(r, state);
      html += piece.html;
      text += piece.text;
    }
    if (!html) classes.push("empty");
    const i = state.texts.length;
    state.texts.push(text);
    const styleAttr = styles.length ? ` style="${styles.join(";")}"` : "";
    return `<p class="${classes.join(" ")}" data-i="${i}"${styleAttr}>${html || "&nbsp;"}</p>`;
  }

  // --------------------------------------------------------- package handling

  function hasContent(doc) {
    return Array.from(documentBody(doc).getElementsByTagNameNS(W, "r")).some(hasVisibleContent);
  }

  function unwrapSdt(doc, tag) {
    if (!tag) return;
    for (const sdt of Array.from(doc.getElementsByTagNameNS(W, "sdt"))) {
      if (sdtTag(sdt) !== tag) continue;
      const content = kid(sdt, "sdtContent");
      while (content && content.firstChild) sdt.parentNode.insertBefore(content.firstChild, sdt);
      sdt.parentNode.removeChild(sdt);
    }
  }

  function slimDoc(doc) {
    const body = documentBody(doc);
    const parts = Array.from(doc.getElementsByTagNameNS(PKG, "part"));
    const usesLists = body && body.getElementsByTagNameNS(W, "numPr").length > 0;
    const dropped = new Set();
    for (const part of parts) {
      const name = part.getAttributeNS(PKG, "name");
      if (DROP_PARTS.some((re) => re.test(name)) || (!usesLists && name === "/word/numbering.xml")) {
        dropped.add(name);
        part.parentNode.removeChild(part);
      }
    }
    for (const part of parts) {
      const name = part.getAttributeNS(PKG, "name");
      if (dropped.has(name) || !/\.rels$/.test(name)) continue;
      const base = name.replace(/_rels\/[^/]*$/, "");
      for (const rel of Array.from(part.getElementsByTagNameNS(REL, "Relationship"))) {
        if (rel.getAttribute("TargetMode") === "External") continue;
        const target = rel.getAttribute("Target");
        const full = target.startsWith("/") ? target : base + target;
        if (dropped.has(full)) rel.parentNode.removeChild(rel);
      }
    }
  }

  // Readies a package for insertOoxml into an inline span: no section break,
  // and a single paragraph contributes only its runs (not its paragraph
  // formatting) so the surrounding paragraph is left alone.
  function forInsertDoc(doc, ownTag) {
    unwrapSdt(doc, ownTag);
    const body = documentBody(doc);
    for (const s of kids(body, "sectPr")) body.removeChild(s);
    const paras = kids(body, "p");
    if (kids(body).length === 1 && paras.length === 1) {
      for (const pPr of kids(paras[0], "pPr")) paras[0].removeChild(pPr);
    }
    slimDoc(doc);
    return serialize(doc);
  }

  function run(xml, ownTag, skipTags, fn) {
    const doc = parse(xml);
    unwrapSdt(doc, ownTag);
    const skip = new Set(skipTags || []);
    const found = new Set();
    const problems = [];
    const changed = fn(doc, skip, found, problems);
    return {
      xml: forInsertDoc(doc, null),
      empty: !hasContent(doc),
      changed,
      found: Array.from(found),
      problems,
    };
  }

  const api = {
    LEVEL_COLOR,
    LEVEL_SHADE,

    /** Mark the selection as struck at `level`. */
    markDelete(xml, level, skipTags, ownTag) {
      return run(xml, ownTag, skipTags, (doc, skip, found) => markDeleteDoc(doc, level, skip, found));
    },

    /** Mark the selection as inserted at `level`. */
    markInsert(xml, level, skipTags, ownTag) {
      return run(xml, ownTag, skipTags, (doc, skip, found) => restyleAllDoc(doc, "ins", level, false, skip, found));
    },

    /** Restyle every run: used for the insertion placeholder and separator. */
    restyleAll(xml, kind, level, gen) {
      return run(xml, null, [], (doc, skip, found) => restyleAllDoc(doc, kind, level, gen, skip, found));
    },

    /** Adopted: fold the amendment at `level` into level-1. */
    fold(xml, level, skipTags, ownTag, tidy) {
      return run(xml, ownTag, skipTags, (doc, skip, found, problems) =>
        foldDoc(doc, level, skip, found, !!tidy, problems));
    },

    /**
     * Adopted/Failed for untracked Level 1 at character `offset` of a single
     * paragraph's OOXML. Result has `blocked` = tag of a pending span that is
     * part of that change (nothing changed), else null.
     */
    resolveUntrackedAt(xml, offset, adopted, pendingTags) {
      const info = {};
      const out = run(xml, null, [], (doc) =>
        resolveUntrackedAtDoc(doc, offset, adopted, new Set(pendingTags || []), info));
      out.blocked = info.blocked || null;
      return out;
    },

    /** Failed, for Level 1 markup with no stored prior text. */
    failUntracked(xml, skipTags) {
      return run(xml, null, skipTags, (doc, skip, found) => failUntrackedDoc(doc, skip, found));
    },

    /**
     * True when the content control tagged `tag` holds no real text: either
     * Word is showing its "Click or tap here to enter text." placeholder
     * (w:showingPlcHdr, or runs in the PlaceholderText style) or it is empty.
     */
    isEmptySpan(xml, tag) {
      const doc = parse(xml);
      let scope = documentBody(doc);
      for (const sdt of Array.from(doc.getElementsByTagNameNS(W, "sdt"))) {
        if (sdtTag(sdt) !== tag) continue;
        if (kid(kid(sdt, "sdtPr"), "showingPlcHdr")) return true;
        scope = kid(sdt, "sdtContent") || sdt;
        break;
      }
      return !Array.from(scope.getElementsByTagNameNS(W, "r")).some((r) =>
        hasVisibleContent(r) && wVal(kid(kid(r, "rPr"), "rStyle")) !== "PlaceholderText");
    },

    /**
     * Presenter view: the passage's paragraphs as HTML, with the add-in's
     * amendment formatting as CSS classes (a-ins1 … a-del3). Each paragraph
     * carries data-i, its position in the passage; texts[i] is its plain text.
     *
     * Note numbers are shown only when `opts.notesBefore` is given: the count
     * of footnotes and endnotes earlier in the document, { footnote, endnote }.
     * `opts.current` ({ kind, num }) highlights that note's reference;
     * `opts.ownNote` ({ kind, num }) numbers a note's own text. `label` is
     * the current or own note's name, such as "Footnote 2". `fmt` is the
     * note number formats found; pass it back as `opts.fmt` to number a
     * note's text (whose package may lack the document's settings) the same.
     */
    presenterView(xml, opts = {}) {
      const doc = parse(xml);
      const fmt = opts.fmt || noteFormats(doc);
      const state = {
        texts: [],
        fmt,
        notes: opts.notesBefore ? { ...opts.notesBefore } : null,
        current: opts.current || null,
        ownNote: opts.ownNote || null,
      };
      const html = presenterBlocks(documentBody(doc), state);
      const note = opts.ownNote || opts.current;
      const label = note
        ? `${note.kind === "endnote" ? "Endnote" : "Footnote"} ${formatNoteNumber(note.num, fmt[note.kind])}`
        : null;
      return { html, texts: state.texts, label, fmt };
    },

    /** Smallest package that still restores the span (stored as priorText). */
    slim(xml) {
      const doc = parse(xml);
      slimDoc(doc);
      return serialize(doc);
    },

    /** Prepare a stored prior package for insertOoxml. */
    forInsert(xml, ownTag) {
      return forInsertDoc(parse(xml), ownTag);
    },

    // Exposed for test/test.html.
    _internal: { classify, runText, restyle, isSep, parse, serialize, documentBody, wEl, setRunText, W, PKG },
  };

  root.AmdOoxml = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof self !== "undefined" ? self : this);
