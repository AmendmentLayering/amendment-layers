/*
 * Tests for web/ooxml.js. Open test.html in a browser.
 *
 * Runs are written as compact tokens:
 *   "text"        plain          "text{+2}"   Level 2 insertion
 *   "text{-1}"    Level 1 strike  "text{-2u1}" Level 1 insertion struck at Level 2
 *   "text{-2c}"   struck copy at Level 2, as Version 2 generated
 *   "·"           generated separator space
 */
(() => {
  "use strict";
  const X = AmdOoxml._internal;
  const results = [];

  function pkg(tokens) {
    const xml =
      `<pkg:package xmlns:pkg="${X.PKG}">` +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${X.W}"><w:body><w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p><w:sectPr/></w:body></w:document>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/xml"><pkg:xmlData><w:styles xmlns:w="${X.W}"/></pkg:xmlData></pkg:part>` +
      `</pkg:package>`;
    const doc = X.parse(xml);
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    appendTokens(doc, p, tokens);
    return X.serialize(doc);
  }

  // A token may also be { sdt: tag, tokens: [...] } for a content control.
  function appendTokens(doc, parent, tokens) {
    for (const tok of tokens) {
      if (typeof tok === "object") {
        const sdt = X.wEl(doc, "sdt");
        const sdtPr = X.wEl(doc, "sdtPr");
        sdtPr.appendChild(X.wEl(doc, "tag", { val: tok.sdt }));
        const content = X.wEl(doc, "sdtContent");
        sdt.appendChild(sdtPr);
        sdt.appendChild(content);
        appendTokens(doc, content, tok.tokens);
        parent.appendChild(sdt);
        continue;
      }
      const r = X.wEl(doc, "r");
      if (tok === "·") {
        X.setRunText(r, " ");
        X.restyle(r, "plain", 0, true);
      } else {
        const m = /^(.*?)(?:\{([+-])(\d)(c?)(?:u(\d))?\})?$/s.exec(tok);
        X.setRunText(r, m[1]);
        if (m[2]) X.restyle(r, m[2] === "+" ? "ins" : "del", Number(m[3]), m[4] === "c", Number(m[5] || 0));
        else X.restyle(r, "plain", 0, false);
      }
      parent.appendChild(r);
    }
  }

  // Character offset of `needle` (plus `into`) in the tokens' plain text.
  function offsetOf(tokens, needle, into) {
    const flat = (ts) => ts.map((t) => (typeof t === "object" ? flat(t.tokens) : t === "·" ? " " : t.replace(/\{.*\}$/s, ""))).join("");
    return flat(tokens).indexOf(needle) + (into || 0);
  }

  function tokens(xml) {
    const doc = X.parse(xml);
    const out = [];
    for (const r of X.documentBody(doc).getElementsByTagNameNS(X.W, "r")) {
      const c = X.classify(r);
      const t = X.runText(r);
      if (X.isSep(r)) out.push("·");
      else if (c.kind === "plain") out.push(t);
      else out.push(`${t}{${c.kind === "ins" ? "+" : "-"}${c.level}${c.gen ? "c" : ""}${c.under ? "u" + c.under : ""}}`);
    }
    // Merge adjacent plain runs so splits don't matter to the assertions.
    return out.reduce((acc, t) => {
      const last = acc[acc.length - 1];
      if (last !== undefined && !/[{·]/.test(last) && !/[{·]/.test(t)) acc[acc.length - 1] = last + t;
      else acc.push(t);
      return acc;
    }, []);
  }

  function test(name, fn) {
    try {
      fn();
      results.push({ name, ok: true });
    } catch (e) {
      results.push({ name, ok: false, msg: e.message });
    }
  }

  function eq(actual, expected) {
    const a = JSON.stringify(actual);
    const e = JSON.stringify(expected);
    if (a !== e) throw new Error(`expected ${e}\n     got ${a}`);
  }

  // ---------------------------------------------------------------- marking

  test("L2 strike of a Level 1 insertion strikes it in place, keeping its underline", () => {
    eq(tokens(AmdOoxml.markDelete(pkg(["$75{+1}"]), 2, []).xml), ["$75{-2u1}"]);
  });

  test("L2 strike of plain text restyles it in place", () => {
    eq(tokens(AmdOoxml.markDelete(pkg(["President"]), 2, []).xml), ["President{-2}"]);
  });

  test("L3 strike of a Level 2 insertion strikes it in place, keeping its underline", () => {
    eq(tokens(AmdOoxml.markDelete(pkg(["$65{+2}"]), 3, []).xml), ["$65{-3u2}"]);
  });

  const sentence = "An accurate reference to a motion’s adoption must not be struck out of the minutes or modified.";

  test("mockup Example 3: a long struck passage appears only once", () => {
    const out = AmdOoxml.markDelete(pkg([`${sentence}{+1}`]), 2, []);
    eq(tokens(out.xml), [`${sentence}{-2u1}`]);
    eq(out.xml.split("struck out of the minutes").length - 1, 1);
  });

  test("mixed selection: adopted words and a lower insertion are each struck once", () => {
    eq(tokens(AmdOoxml.markDelete(pkg(["dues of ", "$75{+1}"]), 2, []).xml), ["dues of {-2}", "$75{-2u1}"]);
  });

  test("strike with spaces on both sides keeps the leading space plain", () => {
    eq(tokens(AmdOoxml.markDelete(pkg([" word "]), 2, []).xml), [" ", "word {-2}"]);
  });

  test("strike with a space on one side strikes it with the words", () => {
    eq(tokens(AmdOoxml.markDelete(pkg([" word"]), 2, []).xml), [" word{-2}"]);
    eq(tokens(AmdOoxml.markDelete(pkg(["word "]), 2, []).xml), ["word {-2}"]);
  });

  test("adopting a strike that had spaces on both sides keeps one space", () => {
    const struck = AmdOoxml.markDelete(pkg([" word "]), 1, []).xml;
    const doc = X.parse(pkg(["the"]));
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    for (const r of X.parse(struck).getElementsByTagNameNS(X.W, "r")) p.appendChild(doc.importNode(r, true));
    const tail = X.wEl(doc, "r");
    X.setRunText(tail, "is");
    p.appendChild(tail);
    eq(tokens(AmdOoxml.fold(X.serialize(doc), 1, [], null, true).xml), ["the is"]);
  });

  test("striking this level's own insertion removes it", () => {
    const out = AmdOoxml.markDelete(pkg(["$65{+2}"]), 2, []);
    eq(out.empty, true);
  });

  test("strike leaves existing deletions alone", () => {
    eq(tokens(AmdOoxml.markDelete(pkg(["$50{-1}", " ", "$75{+1}"]), 2, []).xml),
      ["$50{-1}", " {-2}", "$75{-2u1}"]);
  });

  test("insert restyles the selection at the level", () => {
    eq(tokens(AmdOoxml.markInsert(pkg(["Secretary"]), 2, []).xml), ["Secretary{+2}"]);
  });

  // Run properties as a compact string, for comparing with the mockup.
  function props(xml) {
    return Array.from(X.parse(xml).getElementsByTagNameNS(X.W, "r")).map((r) => {
      const rPr = r.getElementsByTagNameNS(X.W, "rPr")[0];
      return Array.from(rPr ? rPr.children : []).map((c) => {
        const a = ["val", "color", "fill"].map((n) => c.getAttributeNS(X.W, n)).filter((v) => v && v !== "clear" && v !== "auto");
        return a.length ? `${c.localName}=${a.join("/")}` : c.localName;
      }).join(",");
    });
  }

  test("formatting matches three_layer_amendment_mockup_revised.docx", () => {
    eq(props(pkg(["a{+1}", "a{-1}", "a{+2}", "a{-2}", "a{-2u1}", "a{+3}", "a{-3}", "a{-3u2}"])), [
      "color=000000,u=single/000000",
      "strike,color=000000",
      "color=000000,u=double/1F5FBF,shd=DCE9FA",
      "dstrike,color=000000,shd=DCE9FA",
      "dstrike,color=000000,u=single/000000,shd=DCE9FA",
      "i,iCs,color=000000,u=dotted/8B1E8B,shd=F5DDF5",
      "i,iCs,strike,color=000000,shd=F5DDF5",
      "i,iCs,strike,color=000000,u=double/1F5FBF,shd=F5DDF5",
    ]);
  });

  test("Version 2's colored text is still read at its level", () => {
    const xml = rawPkg(
      `<w:r><w:rPr><w:color w:val="1F5FBF"/><w:u w:val="double"/><w:shd w:val="clear" w:color="auto" w:fill="DCE9FA"/></w:rPr><w:t>a</w:t></w:r>` +
      `<w:r><w:rPr><w:i/><w:strike/><w:color w:val="8B1E8B"/><w:shd w:val="clear" w:color="auto" w:fill="F5DDF5"/></w:rPr><w:t>b</w:t></w:r>`);
    eq(tokens(xml), ["a{+2}", "b{-3}"]);
  });

  // ---------------------------------------------------------------- adopted

  test("mockup Example 1: L3 adopted removes the struck $65 and makes $60 the primary's figure", () => {
    const span = ["$65{-3u2}", "·", "$60{+3}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 3, []).xml), ["$60{+2}"]);
  });

  test("then L2 adopted removes the struck $75 and makes $60 the main motion's figure", () => {
    const span = ["$75{-2u1}", "·", "$60{+2}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 2, []).xml), ["$60{+1}"]);
  });

  test("mockup Example 3: L3 adopted inside the primary's replacement sentence", () => {
    const span = ["in the mistaken belief that doing so will{-3u2}", "·", "; that will not{+3}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 3, []).xml), ["; that will not{+2}"]);
  });

  test("L3 strike of a Level 1 insertion becomes a Level 2 strike when adopted, then goes", () => {
    const l3 = AmdOoxml.fold(pkg(["monthly{-3u1}", "·", "weekly{+3}"]), 3, []).xml;
    eq(tokens(l3), ["monthly{-2u1}", "·", "weekly{+2}"]);
    eq(tokens(AmdOoxml.fold(l3, 2, []).xml), ["weekly{+1}"]);
  });

  // Documents with amendments still pending from Version 2 have struck copies.
  test("Version 2 struck copies: L3 adopted", () => {
    const span = ["$65{+2}", "·", "$65{-3c}", "·", "$60{+3}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 3, []).xml), ["$60{+2}"]);
  });

  test("Version 2 struck copies: L2 adopted", () => {
    const span = ["$75{+1}", "·", "$75{-2c}", "·", "$60{+2}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 2, []).xml), ["$60{+1}"]);
  });

  test("mockup Example 2: L2 adopted moves in-place strike and insertion to Level 1", () => {
    const span = ["President{-2}", "·", "Secretary{+2}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 2, []).xml), ["President{-1}", "·", "Secretary{+1}"]);
  });

  test("tracked L1 adopted leaves plain final text", () => {
    eq(tokens(AmdOoxml.fold(pkg(["President{-1}", "·", "Secretary{+1}"]), 1, []).xml), ["Secretary"]);
  });

  test("L3 strike of a Level 1 insertion becomes a Level 2 strike when adopted", () => {
    const span = ["monthly{+1}", "·", "monthly{-3c}", "·", "weekly{+3}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 3, []).xml), ["monthly{+1}", "·", "monthly{-2c}", "·", "weekly{+2}"]);
  });

  test("copy matching a part of a longer insertion splits the original", () => {
    const span = ["$65 annually{+2}", "·", "annually{-3c}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 3, []).xml), ["$65 {+2}"]);
  });

  // From the Bylaw 3 test document: the Level 3 selection caught the Level 1
  // space after "1.0%".
  test("strike leaves a lower-level space at the selection's edge alone", () => {
    eq(tokens(AmdOoxml.markDelete(pkg(["1.0%{+2}", " {+1}"]), 3, []).xml),
      ["1.0%{-3u2}", " {+1}"]);
  });

  test("a copy with a trailing space still matches its original", () => {
    const out = AmdOoxml.fold(pkg(["1.0%{+2}", "·", "1.0% {-3c}", "·", "2.0%{+3}"]), 3, []);
    eq(out.problems, []);
    eq(tokens(out.xml), ["2.0%{+2}"]);
  });

  test("an original with a trailing space still matches its copy", () => {
    const out = AmdOoxml.fold(pkg(["1.0% {+2}", "·", "1.0%{-3c}"]), 3, []);
    eq(out.problems, []);
    eq(out.empty, true);
  });

  test("a whitespace-only copy is simply dropped", () => {
    const out = AmdOoxml.fold(pkg(["x{+2}", " {+1}", "·", " {-3c}"]), 3, []);
    eq(out.problems, []);
    eq(tokens(out.xml), ["x{+2}", " {+1}"]);
  });

  test("a struck copy with no matching original is flagged, not guessed", () => {
    const out = AmdOoxml.fold(pkg(["$65{-3c}"]), 3, []);
    eq(out.problems.length, 1);
  });

  test("lower and higher levels are untouched by a fold", () => {
    const span = ["$50{-1}", " ", "x{+3}", "$65{-2}"];
    eq(tokens(AmdOoxml.fold(pkg(span), 2, []).xml), ["$50{-1}", " ", "x{+3}", "$65{-1}"]);
  });

  // ----------------------------------------------------- untracked Level 1

  const noticed = ["The annual dues shall be ", "$50{-1}", " ", "$75{+1}", ", payable by ", "January 31{-1}", " ", "March 1{+1}", "."];

  test("untracked Level 1 adopted", () => {
    eq(tokens(AmdOoxml.fold(pkg(noticed), 1, [], null, true).xml),
      ["The annual dues shall be $75, payable by March 1."]);
  });

  test("untracked Level 1 failed", () => {
    eq(tokens(AmdOoxml.failUntracked(pkg(noticed), []).xml),
      ["The annual dues shall be $50, payable by January 31."]);
  });

  // ------------------------------------------ untracked Level 1 at the cursor

  test("cursor in one Level 1 change adopts only that change", () => {
    const out = AmdOoxml.resolveUntrackedAt(pkg(noticed), offsetOf(noticed, "$50", 1), true, []);
    eq(out.blocked, null);
    eq(tokens(out.xml), ["The annual dues shall be $75, payable by ", "January 31{-1}", " ", "March 1{+1}", "."]);
  });

  test("cursor in the other Level 1 change fails only that change", () => {
    const out = AmdOoxml.resolveUntrackedAt(pkg(noticed), offsetOf(noticed, "March", 2), false, []);
    eq(tokens(out.xml), ["The annual dues shall be ", "$50{-1}", " ", "$75{+1}", ", payable by January 31."]);
  });

  test("cursor right after the change's last word still finds it", () => {
    const out = AmdOoxml.resolveUntrackedAt(pkg(noticed), offsetOf(noticed, "March 1", 7), true, []);
    eq(out.changed, 2);
  });

  const commaLeft = ["Whether to (a) extend the date", ", (b) modify the requirements{-1}", ", ",
    "(b) rescind the amendment,{+1}", " or (c) present a resolution."];

  test("a comma left unstruck between a Level 1 strike and insertion keeps them one change", () => {
    const out = AmdOoxml.resolveUntrackedAt(pkg(commaLeft), offsetOf(commaLeft, "modify", 1), true, []);
    eq(out.changed, 2);
    eq(tokens(out.xml), ["Whether to (a) extend the date, (b) rescind the amendment, or (c) present a resolution."]);
  });

  test("words between Level 1 changes still keep them apart", () => {
    const apart = ["Dues ", "$50{-1}", ", payable ", "$75{+1}", "."];
    eq(AmdOoxml.resolveUntrackedAt(pkg(apart), offsetOf(apart, "$50", 1), true, []).changed, 1);
  });

  test("cursor in plain words finds no Level 1 change", () => {
    eq(AmdOoxml.resolveUntrackedAt(pkg(noticed), offsetOf(noticed, "annual", 2), true, []).changed, 0);
  });

  const withPrimary = ["Dues are ", "$50{-1}", " ", { sdt: "amd-2-1-x", tokens: ["$75{-2u1}"] },
    " ", "$65{+2}", ", payable by ", "January 31{-1}", " ", "March 1{+1}", "."];

  test("order guard: a pending primary amendment to this change blocks it", () => {
    const out = AmdOoxml.resolveUntrackedAt(pkg(withPrimary), offsetOf(withPrimary, "$50", 1), true, ["amd-2-1-x"]);
    eq(out.blocked, "amd-2-1-x");
    eq(out.changed, 0);
  });

  test("order guard: a pending amendment to another change in the paragraph doesn't block", () => {
    const out = AmdOoxml.resolveUntrackedAt(pkg(withPrimary), offsetOf(withPrimary, "March", 1), true, ["amd-2-1-x"]);
    eq(out.blocked, null);
    eq(out.changed, 2);
  });

  // ---------------------------------------------------------- presenter view

  test("presenter view shows each level with its class", () => {
    const view = AmdOoxml.presenterView(pkg(["Dues ", "$50{-1}", " ", "$75{+1}", " ", "$65{+3}"]));
    for (const cls of ["a-del1", "a-ins1", "a-ins3"]) {
      if (!view.html.includes(`class="${cls}"`)) throw new Error(`missing ${cls}`);
    }
    eq(view.texts, ["Dues $50 $75 $65"]);
  });

  test("presenter view shows a struck insertion once, keeping its underline", () => {
    const view = AmdOoxml.presenterView(pkg(["$75{-2u1}", " ", "$65{-3u2}"]));
    if (!view.html.includes(`<span class="a-u1"><span class="a-del2">$75</span></span>`)) throw new Error(view.html);
    if (!view.html.includes(`<span class="a-u2"><span class="a-del3">$65</span></span>`)) throw new Error(view.html);
    eq(view.texts, ["$75 $65"]);
  });

  test("presenter view numbers paragraphs in Word's order, table cells included", () => {
    const xml = `<pkg:package xmlns:pkg="${X.PKG}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="x"><pkg:xmlData>` +
      `<w:document xmlns:w="${X.W}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p>` +
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>three</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>four</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const view = AmdOoxml.presenterView(xml);
    eq(view.texts, ["one", "two", "three", "four"]);
    if (!/class="p h2" data-i="3"/.test(view.html)) throw new Error("heading not marked");
    if (!/<table class="t">/.test(view.html)) throw new Error("table missing");
  });

  test("presenter view skips Word's placeholder text and escapes HTML", () => {
    const xml = withSdt("amd-1-2-x", "<w:showingPlcHdr/>",
      `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r><w:r><w:t>a &lt; b</w:t></w:r>`);
    const view = AmdOoxml.presenterView(xml);
    if (/Click or tap/.test(view.html)) throw new Error("placeholder shown");
    if (!/a &lt; b/.test(view.html)) throw new Error("not escaped");
  });

  // A one-paragraph package from raw runs, with optional settings.xml content.
  function rawPkg(runsXml, settings) {
    return `<pkg:package xmlns:pkg="${X.PKG}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="x"><pkg:xmlData>` +
      `<w:document xmlns:w="${X.W}"><w:body><w:p>${runsXml}</w:p></w:body></w:document></pkg:xmlData></pkg:part>` +
      (settings ? `<pkg:part pkg:name="/word/settings.xml" pkg:contentType="x"><pkg:xmlData><w:settings xmlns:w="${X.W}">${settings}</w:settings></pkg:xmlData></pkg:part>` : "") +
      `</pkg:package>`;
  }
  const fnRef = (id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  const enRef = (id) => `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
  const text = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

  test("presenter view numbers notes after those earlier in the document", () => {
    const xml = rawPkg(text("Nine directors") + fnRef(7) + text(" for three years") + fnRef(9) + enRef(4));
    const view = AmdOoxml.presenterView(xml, { notesBefore: { footnote: 2, endnote: 0 } });
    const nums = [...view.html.matchAll(/<sup class="nref">([^<]*)<\/sup>/g)].map((m) => m[1]);
    eq(nums, ["3", "4", "i"]); // Word's default endnote numbers are roman
    eq(view.texts, ["Nine directors for three years"]);
  });

  test("presenter view highlights the current note and names it", () => {
    const xml = rawPkg(text("Terms") + fnRef(1) + text(" staggered") + fnRef(2), "<w:footnotePr><w:numFmt w:val=\"lowerLetter\"/></w:footnotePr>");
    const view = AmdOoxml.presenterView(xml, { notesBefore: { footnote: 0, endnote: 0 }, current: { kind: "footnote", num: 2 } });
    if (!/<sup class="nref">a<\/sup>/.test(view.html)) throw new Error("first note not plain");
    if (!/<sup class="nref cur">b<\/sup>/.test(view.html)) throw new Error("current note not highlighted");
    eq(view.label, "Footnote b");
    eq(view.fmt.footnote, "lowerLetter");
  });

  test("presenter view numbers a note's own text", () => {
    const xml = rawPkg(`<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>` + text(" Transition proviso."));
    const view = AmdOoxml.presenterView(xml, { notesBefore: { footnote: 0, endnote: 0 }, ownNote: { kind: "endnote", num: 1 }, fmt: { footnote: "decimal", endnote: "decimal" } });
    if (!/<sup class="nref">1<\/sup> Transition proviso\./.test(view.html)) throw new Error(view.html);
    eq(view.label, "Endnote 1");
  });

  test("presenter view leaves note numbers out unless asked", () => {
    const view = AmdOoxml.presenterView(rawPkg(text("Terms") + fnRef(1)));
    if (/nref/.test(view.html)) throw new Error("note number shown");
  });

  test("presenter view gives a custom note mark no number", () => {
    const xml = rawPkg(text("Terms") + `<w:r><w:footnoteReference w:customMarkFollows="1" w:id="1"/><w:t>*</w:t></w:r>` + fnRef(2));
    const view = AmdOoxml.presenterView(xml, { notesBefore: { footnote: 0, endnote: 0 } });
    const nums = [...view.html.matchAll(/<sup class="nref">([^<]*)<\/sup>/g)].map((m) => m[1]);
    eq(nums, ["1"]);
  });

  // ------------------------------------------------------------ list numbers

  // A bylaws-style list: Article I, Section 1., (a). Each paragraph is
  // [level, tokens] (numPr on the paragraph), ["H", tokens] (Heading 1, numbered
  // through its style), or [null, tokens] (no number). `extra` adds w:num
  // entries; a paragraph [level, tokens, numId] uses one.
  function listPkg(paras, extra) {
    const doc = X.parse(
      `<pkg:package xmlns:pkg="${X.PKG}">` +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="x"><pkg:xmlData><w:document xmlns:w="${X.W}"><w:body/></w:document></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="x"><pkg:xmlData><w:numbering xmlns:w="${X.W}">` +
      `<w:abstractNum w:abstractNumId="0">` +
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Article %1"/><w:pStyle w:val="Heading1"/></w:lvl>` +
      `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Section %2."/></w:lvl>` +
      `<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/></w:lvl>` +
      `</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
      `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>` + (extra || "") +
      `</w:numbering></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="x"><pkg:xmlData><w:styles xmlns:w="${X.W}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>` +
      `</w:styles></pkg:xmlData></pkg:part></pkg:package>`);
    const body = X.documentBody(doc);
    for (const [lvl, toks, numId] of paras) {
      const p = X.wEl(doc, "p");
      const pPr = X.wEl(doc, "pPr");
      if (lvl === "H") {
        pPr.appendChild(X.wEl(doc, "pStyle", { val: "Heading1" }));
      } else if (lvl != null) {
        const np = X.wEl(doc, "numPr");
        np.appendChild(X.wEl(doc, "ilvl", { val: String(lvl) }));
        np.appendChild(X.wEl(doc, "numId", { val: String(numId || 1) }));
        pPr.appendChild(np);
      }
      p.appendChild(pPr);
      appendTokens(doc, p, toks);
      body.appendChild(p);
    }
    return X.serialize(doc);
  }
  const field = (list, key) => list.map((n) => (n ? n[key] : null));

  test("list numbers: Word's numbering, styles and restarts included", () => {
    const list = AmdOoxml.listNumbers(listPkg([
      ["H", ["Name"]], [1, ["Name of club"]], ["H", ["Members"]], [1, ["Classes"]],
      [2, ["Active"]], [2, ["Honorary"]], [1, ["Dues"]], [null, ["Note"]], [2, ["Annual"]],
    ]));
    eq(field(list, "word"), ["Article I", "Section 1.", "Article II", "Section 1.", "(a)", "(b)", "Section 2.", null, "(a)"]);
  });

  test("list numbers: nothing pending, before and after match Word's", () => {
    const list = AmdOoxml.listNumbers(listPkg([[1, ["One"]], [1, ["Two"]]]));
    eq(field(list, "orig"), ["Section 1.", "Section 2."]);
    eq(field(list, "prop"), ["Section 1.", "Section 2."]);
    eq(field(list, "status"), [null, null]);
  });

  test("list numbers: a wholly struck subsection renumbers those after it", () => {
    const list = AmdOoxml.listNumbers(listPkg([
      [2, ["Active"]], [2, ["Honorary{-2}"]], [2, ["Life"]], [2, ["Student"]], [1, ["Dues"]], [2, ["Annual"]],
    ]));
    eq(field(list, "word"), ["(a)", "(b)", "(c)", "(d)", "Section 1.", "(a)"]);
    eq(field(list, "orig"), ["(a)", "(b)", "(c)", "(d)", "Section 1.", "(a)"]);
    eq(field(list, "prop"), ["(a)", null, "(b)", "(c)", "Section 1.", "(a)"]);
    eq(field(list, "status"), [null, "del", null, null, null, null]);
    eq(field(list, "cause"), [0, 0, 2, 2, 0, 0]);
  });

  test("list numbers: a wholly inserted section renumbers those after it", () => {
    const list = AmdOoxml.listNumbers(listPkg([[1, ["One"]], [1, ["New{+1}"]], [1, ["Two"]]]));
    eq(field(list, "word"), ["Section 1.", "Section 2.", "Section 3."]);
    eq(field(list, "orig"), ["Section 1.", null, "Section 2."]);
    eq(field(list, "prop"), ["Section 1.", "Section 2.", "Section 3."]);
    eq(field(list, "status"), [null, "ins", null]);
    eq(field(list, "level"), [0, 1, 0]);
  });

  test("list numbers: a partly struck paragraph keeps its number", () => {
    const list = AmdOoxml.listNumbers(listPkg([[1, ["One ", "two{-2}"]], [1, ["Three"]]]));
    eq(field(list, "status"), [null, null]);
    eq(field(list, "prop"), ["Section 1.", "Section 2."]);
  });

  test("list numbers: an insertion wholly struck counts neither before nor after", () => {
    const list = AmdOoxml.listNumbers(listPkg([[1, ["One"]], [1, ["New{-2u1}"]], [1, ["Two"]]]));
    eq(field(list, "status"), [null, "gone", null]);
    eq(field(list, "orig"), ["Section 1.", null, "Section 2."]);
    eq(field(list, "prop"), ["Section 1.", null, "Section 2."]);
  });

  test("list numbers: a startOverride restarts where its list is first used", () => {
    const extra = `<w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`;
    const list = AmdOoxml.listNumbers(listPkg([[1, ["A"]], [1, ["B"]], [1, ["C"], 5], [1, ["D"], 5]], extra));
    eq(field(list, "word"), ["Section 1.", "Section 2.", "Section 1.", "Section 2."]);
  });

  test("list numbers: bullets are shown as a dot", () => {
    const list = AmdOoxml.listNumbers(listPkg([[0, ["Item"], 2]]));
    eq(field(list, "word"), ["•"]);
  });

  test("presenter view draws numbers: changed ones old struck, new inserted", () => {
    const xml = listPkg([[2, ["Honorary{-2}"]], [2, ["Life"]], [2, ["Student"]]]);
    const nums = AmdOoxml.listNumbers(xml).map((n, i) => ({ ...n, text: ["(a)", "(b)", "(c)"][i] }));
    nums[2] = { ...nums[2], orig: "(c)", prop: "(c)" }; // unchanged number
    const view = AmdOoxml.presenterView(xml, { numbers: nums });
    const spans = [...view.html.matchAll(/<span class="num( [^"]*)?">([^<]*)<\/span>/g)].map((m) => `${m[2]}${m[1] || ""}`);
    eq(spans, ["(a) a-del2", "(b) a-del2", "(a) a-ins2", "(c)"]);
    eq(view.texts, ["Honorary", "Life", "Student"]);
  });

  test("presenter view draws Word's number alone when the arithmetic is unchecked", () => {
    const xml = listPkg([[1, ["One"]]]);
    const view = AmdOoxml.presenterView(xml, { numbers: [{ text: "Section 4." }] });
    if (!/<span class="num">Section 4\.<\/span>/.test(view.html)) throw new Error(view.html);
  });

  // ------------------------------------------------------------ packaging

  test("skip tags leave another pending amendment's span alone", () => {
    const doc = X.parse(pkg(["$75{+1}"]));
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    const sdt = X.wEl(doc, "sdt");
    const sdtPr = X.wEl(doc, "sdtPr");
    sdtPr.appendChild(X.wEl(doc, "tag", { val: "amd-9-1-x" }));
    const content = X.wEl(doc, "sdtContent");
    const r = X.wEl(doc, "r");
    X.setRunText(r, "inner");
    content.appendChild(r);
    sdt.appendChild(sdtPr);
    sdt.appendChild(content);
    p.appendChild(sdt);
    const out = AmdOoxml.markDelete(X.serialize(doc), 2, ["amd-9-1-x"]);
    eq(tokens(out.xml), ["$75{-2u1}", "inner"]);
    eq(out.found, ["amd-9-1-x"]);
  });

  test("own content control wrapper is unwrapped before transforming", () => {
    const doc = X.parse(pkg([]));
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    const sdt = X.wEl(doc, "sdt");
    const sdtPr = X.wEl(doc, "sdtPr");
    sdtPr.appendChild(X.wEl(doc, "tag", { val: "amd-1-1-x" }));
    const content = X.wEl(doc, "sdtContent");
    const r = X.wEl(doc, "r");
    X.setRunText(r, "words");
    content.appendChild(r);
    sdt.appendChild(sdtPr);
    sdt.appendChild(content);
    p.appendChild(sdt);
    const out = AmdOoxml.markDelete(X.serialize(doc), 2, [], "amd-1-1-x");
    if (/w:sdt/.test(out.xml)) throw new Error("sdt still present");
    eq(tokens(out.xml), ["words{-2}"]);
  });

  test("insert-ready package drops styles part, sectPr and single-paragraph pPr", () => {
    const xml = AmdOoxml.markInsert(pkg(["a"]), 2, []).xml;
    if (/styles\.xml/.test(xml)) throw new Error("styles part kept");
    if (/w:sectPr|w:pPr/.test(xml)) throw new Error("sectPr/pPr kept");
  });

  function withSdt(tag, sdtPrExtra, runsXml) {
    return `<pkg:package xmlns:pkg="${X.PKG}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${X.W}"><w:body><w:p><w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${sdtPrExtra}</w:sdtPr>` +
      `<w:sdtContent>${runsXml}</w:sdtContent></w:sdt></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  }

  test("a span showing Word's placeholder counts as empty", () => {
    const xml = withSdt("amd-1-2-x", "<w:showingPlcHdr/>",
      `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r>`);
    eq(AmdOoxml.isEmptySpan(xml, "amd-1-2-x"), true);
  });

  test("placeholder-styled runs alone count as empty", () => {
    const xml = withSdt("amd-1-2-x", "",
      `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r>`);
    eq(AmdOoxml.isEmptySpan(xml, "amd-1-2-x"), true);
  });

  test("a span with real text is not empty", () => {
    eq(AmdOoxml.isEmptySpan(withSdt("amd-1-2-x", "", `<w:r><w:t>2.0%</w:t></w:r>`), "amd-1-2-x"), false);
  });

  test("an empty run list is empty", () => {
    eq(AmdOoxml.isEmptySpan(withSdt("amd-1-2-x", "", ""), "amd-1-2-x"), true);
  });

  // ---------------------------------------------------------- paragraph marks

  // A paragraph whose end mark (and so its list number) is struck at `level`.
  function withMark(tokens, level) {
    const doc = X.parse(pkg(tokens));
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    const r = X.wEl(doc, "r");
    X.restyle(r, level ? "del" : "plain", level, false);
    p.getElementsByTagNameNS(X.W, "pPr")[0].appendChild(r.getElementsByTagNameNS(X.W, "rPr")[0]);
    return X.serialize(doc);
  }

  test("a paragraph struck end to end wants its number struck at that level", () => {
    eq(AmdOoxml.markInfo(pkg(["Presentation of Awards. No award shall be presented.{-2}"])),
      { mark: 0, want: 2, empty: false });
  });

  test("a strike inside an add-in span counts toward the whole paragraph", () => {
    eq(AmdOoxml.markInfo(pkg([{ sdt: "amd-1-1-x", tokens: ["Presentation of Awards.{-2}", " No award.{-2}"] }])).want, 2);
  });

  test("a partly struck paragraph leaves its number alone", () => {
    eq(AmdOoxml.markInfo(pkg(["Presentation of Awards.", " No award.{-2}"])).want, 0);
  });

  test("spaces left plain at the edge don't stop the number being struck", () => {
    eq(AmdOoxml.markInfo(pkg(["No award.{-1}", " "])).want, 1);
  });

  test("mixed levels strike the number at the highest one", () => {
    eq(AmdOoxml.markInfo(pkg(["Awards.{-1}", " No award.{-2}"])).want, 2);
  });

  test("a visible insertion keeps the number plain; a struck one counts as struck", () => {
    eq(AmdOoxml.markInfo(pkg(["$75{+1}", " No award.{-2}"])).want, 0);
    eq(AmdOoxml.markInfo(pkg(["$75{-2u1}", " No award.{-2}"])).want, 2);
  });

  test("a struck number with plain text again wants to go back to plain", () => {
    eq(AmdOoxml.markInfo(withMark(["No award shall be presented."], 2)), { mark: 2, want: 0, empty: false });
  });

  test("a Level 2 number becomes Level 1 when its strike is adopted", () => {
    eq(AmdOoxml.markInfo(withMark(["No award.{-1}"], 2)), { mark: 2, want: 1, empty: false });
  });

  test("an emptied paragraph keeps its struck number, to be removed", () => {
    eq(AmdOoxml.markInfo(withMark([], 1)), { mark: 1, want: 1, empty: true });
    eq(AmdOoxml.markInfo(withMark(["​"], 1)).empty, true);
  });

  test("a Level 1 strike of a whole paragraph, adopted, leaves it empty", () => {
    const struck = AmdOoxml.markDelete(pkg(["Presentation of Awards. No award shall be presented."]), 1, []).xml;
    eq(AmdOoxml.fold(struck, 1, [], null, true).empty, true);
  });

  // The paragraph wrapped in a block-level span, as Word makes it when a span
  // covers all of a paragraph's text (seen in saved test documents).
  function inBlockSpan(xml, tag) {
    const doc = X.parse(xml);
    const body = X.documentBody(doc);
    const p = body.getElementsByTagNameNS(X.W, "p")[0];
    const sdt = X.wEl(doc, "sdt");
    const sdtPr = X.wEl(doc, "sdtPr");
    sdtPr.appendChild(X.wEl(doc, "tag", { val: tag }));
    const content = X.wEl(doc, "sdtContent");
    sdt.appendChild(sdtPr);
    sdt.appendChild(content);
    body.insertBefore(sdt, p);
    content.appendChild(p);
    return X.serialize(doc);
  }

  const awards = "Presentation of Awards. No award shall be presented.";

  test("before a strike: the number is predicted struck from the span's new contents", () => {
    const para = pkg([{ sdt: "amd-1-1-x", tokens: [awards] }]);
    const struck = AmdOoxml.markDelete(pkg([awards]), 2, []).xml;
    eq(AmdOoxml.markInfoAfter(para, { "amd-1-1-x": struck }), { mark: 0, want: 2, empty: false });
  });

  test("before a strike of part of a paragraph: the number stays plain", () => {
    const para = pkg(["Presentation of Awards.", { sdt: "amd-1-1-x", tokens: [" No award."] }]);
    const struck = AmdOoxml.markDelete(pkg([" No award."]), 2, []).xml;
    eq(AmdOoxml.markInfoAfter(para, { "amd-1-1-x": struck }).want, 0);
  });

  test("before adopting a block-level Level 2 strike: the number becomes Level 1", () => {
    const struck = AmdOoxml.markDelete(pkg([awards]), 2, []).xml;
    const para = inBlockSpan(withMark([`${awards}{-2}`], 2), "amd-1-1-x");
    const folded = AmdOoxml.fold(struck, 2, [], null, false).xml;
    eq(AmdOoxml.markInfoAfter(para, { "amd-1-1-x": folded }), { mark: 2, want: 1, empty: false });
  });

  test("before failing a block-level strike: the number goes back to plain", () => {
    const para = inBlockSpan(withMark([`${awards}{-2}`], 2), "amd-1-1-x");
    eq(AmdOoxml.markInfoAfter(para, { "amd-1-1-x": pkg([awards]) }), { mark: 2, want: 0, empty: false });
  });

  test("a span not in the paragraph's package replaces all its runs", () => {
    const para = withMark([`${awards}{-2}`], 2);
    eq(AmdOoxml.markInfoAfter(para, { "amd-9-9-x": pkg([awards]) }).want, 0);
  });

  test("an emptied span leaves the rest of the paragraph deciding", () => {
    const para = pkg(["Kept words.{-1}", { sdt: "amd-1-1-x", tokens: [" new words{+2}"] }]);
    eq(AmdOoxml.markInfoAfter(para, { "amd-1-1-x": null }).want, 1);
  });

  test("no replacements: the paragraph as it is", () => {
    eq(AmdOoxml.markInfoAfter(pkg([`${awards}{-2}`]), {}).want, 2);
  });

  test("a Level 2 strike of a whole paragraph, adopted, becomes a Level 1 strike", () => {
    const struck = AmdOoxml.markDelete(pkg(["No award shall be presented."]), 2, []).xml;
    eq(AmdOoxml.markInfo(AmdOoxml.fold(struck, 2, [], null, false).xml).want, 1);
  });

  test("rPr children stay in schema order", () => {
    const xml = AmdOoxml.markDelete(pkg(["a{+1}"]), 3, []).xml;
    const doc = X.parse(xml);
    const order = ["i", "iCs", "strike", "dstrike", "noProof", "color", "u", "shd"];
    for (const rPr of doc.getElementsByTagNameNS(X.W, "rPr")) {
      const names = Array.from(rPr.children).map((c) => c.localName);
      const idx = names.map((n) => order.indexOf(n));
      if (idx.some((v, i) => i && v < idx[i - 1])) throw new Error(`bad order: ${names.join(",")}`);
    }
  });

  // ------------------------------------------- manual markup -> Track Changes

  // [tag, text] for each conversion control, in order; plain text as is.
  function tcPieces(xml) {
    const doc = X.parse(xml);
    const out = [];
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    for (const c of Array.from(p.children)) {
      if (c.localName === "sdt") {
        const tag = c.getElementsByTagNameNS(X.W, "tag")[0].getAttributeNS(X.W, "val");
        const runs = Array.from(c.getElementsByTagNameNS(X.W, "r"));
        for (const r of runs) {
          const rPr = r.getElementsByTagNameNS(X.W, "rPr")[0];
          if (rPr && rPr.getElementsByTagNameNS(X.W, tag.includes("del") ? "strike" : "u").length) {
            throw new Error(`${tag} still has its markup`);
          }
        }
        out.push([tag, runs.map(X.runText).join("")]);
      } else if (c.localName === "r") {
        out.push(X.runText(c));
      }
    }
    return out;
  }

  test("markup to Track Changes: struck and underlined stretches get their own controls", () => {
    const out = AmdOoxml.markupToTracked(pkg(["The Board shall meet ", "quarterly{-1}", " ", "monthly{+1}", " at noon."]), "t-", { n: 1 });
    eq(out.changed, true);
    eq(tcPieces(out.xml), ["The Board shall meet ", ["t-del-1", "quarterly"], " ", ["t-ins-2", "monthly"], " at noon."]);
  });

  test("markup to Track Changes: adjacent runs of one kind share a control", () => {
    const counter = { n: 5 };
    const out = AmdOoxml.markupToTracked(pkg(["a", "in 3{-1}", " years{-1}", "b"]), "t-", counter);
    eq(tcPieces(out.xml), ["a", ["t-del-5", "in 3 years"], "b"]);
    eq(counter.n, 6);
  });

  test("markup to Track Changes: a note's own number is left out (Word keeps its own)", () => {
    const doc = X.parse(pkg([" Any ", "action taken{-1}", " rest."]));
    const p = X.documentBody(doc).getElementsByTagNameNS(X.W, "p")[0];
    const r = doc.createElementNS(X.W, "w:r");
    r.appendChild(doc.createElementNS(X.W, "w:footnoteRef"));
    p.insertBefore(r, p.getElementsByTagNameNS(X.W, "r")[0]);
    const out = AmdOoxml.markupToTracked(X.serialize(doc), "t-", { n: 1 });
    eq(out.changed, true);
    eq(X.parse(out.xml).getElementsByTagNameNS(X.W, "footnoteRef").length, 0);
    // The space after the number goes too: Word keeps it with the number.
    eq(tcPieces(out.xml), ["Any ", ["t-del-1", "action taken"], " rest."]);
  });

  test("markup to Track Changes: notes carried in a main-text paragraph's package keep their numbers", () => {
    // A main-text paragraph's package also holds the notes it refers to;
    // Word rebuilds them from it, so their numbers must stay.
    const doc = X.parse(pkg(["Law ", "null and void{-1}", " applies."]));
    const part = doc.createElementNS(X.PKG, "pkg:part");
    part.setAttributeNS(X.PKG, "pkg:name", "/word/footnotes.xml");
    part.innerHTML = `<pkg:xmlData xmlns:pkg="${X.PKG}"><w:footnotes xmlns:w="${X.W}"><w:footnote w:id="2"><w:p>` +
      `<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>` +
      `<w:r><w:t xml:space="preserve"> If the assembly is a lawmaking body.</w:t></w:r></w:p></w:footnote></w:footnotes></pkg:xmlData>`;
    doc.documentElement.appendChild(part);
    const out = AmdOoxml.markupToTracked(X.serialize(doc), "t-", { n: 1 });
    eq(out.changed, true);
    const res = X.parse(out.xml);
    eq(res.getElementsByTagNameNS(X.W, "footnoteRef").length, 1);
    eq(res.getElementsByTagNameNS(X.W, "pStyle").length, 1);
    eq(tcPieces(out.xml), ["Law ", ["t-del-1", "null and void"], " applies."]);
  });

  test("markup to Track Changes: Level 2 and 3 formatting is not converted", () => {
    const out = AmdOoxml.markupToTracked(pkg(["x{+2}", "y{-2}", "z{+3}"]), "t-", { n: 1 });
    eq(out.changed, false);
  });

  // ---------------------------------------------------------------- report

  const failed = results.filter((r) => !r.ok);
  document.getElementById("summary").textContent =
    `${results.length - failed.length} / ${results.length} passed`;
  document.getElementById("summary").className = failed.length ? "fail" : "pass";
  document.getElementById("out").innerHTML = results
    .map((r) => `<div class="${r.ok ? "pass" : "fail"}">${r.ok ? "✓" : "✗"} ${r.name}</div>` +
      (r.ok ? "" : `<pre>${r.msg.replace(/</g, "&lt;")}</pre>`))
    .join("");
  window.testResults = results;
})();
