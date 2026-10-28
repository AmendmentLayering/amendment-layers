/*
 * Word-side engine: thread state, anchors (content controls), and the five
 * actions from the spec (set level, strike, insert, adopted, failed).
 *
 * Thread state lives in the document's add-in settings, so a pending thread
 * survives closing and reopening the file (once the file is saved).
 *
 * Thread {
 *   id, seq            seq orders threads; resolve picks the highest seq at the cursor
 *   level              1 | 2 | 3
 *   parentId           the pending level-1 thread this one amends (null for Level 1)
 *   status             pending | adopted | failed | mixed | withdrawn
 *   anchors            pending spans: [{ tag, prior, place, parentPlace }]
 *                      tag          hidden rich-text content control around the span
 *                      prior        the span's OOXML before this thread touched it
 *                                   ("" for an insertion, typed at the cursor or
 *                                   typed first and then marked)
 *                      place        which location of the amendment the span is in
 *                                   (an amendment moved in several places has one
 *                                   place per location; a strike and its
 *                                   replacement words share a place)
 *                      parentPlace  the place of the amended amendment that this
 *                                   span amends; the order guard follows this link
 *   resolved           [{ tag, outcome }] for places already voted
 * }
 *
 * Adopted/Failed act on the place at the cursor; the pane's "All places"
 * button makes the next one act on every place of that amendment.
 */
const Engine = (() => {
  "use strict";

  const STATE_KEY = "amendmentLayers";
  const TAG_PREFIX = "amd-";
  const MAX_VOTES = 50; // votes remembered so Ctrl+Z on them can be followed
  const LEVEL_NAMES = { 1: "Level 1 (main motion)", 2: "Level 2 (primary amendment)", 3: "Level 3 (secondary amendment)" };

  const INSIDE = new Set(["Inside", "InsideStart", "InsideEnd", "Equal"]);
  const AT_END = new Set(["InsideEnd", "AdjacentAfter"]);
  const AT_START = new Set(["InsideStart", "AdjacentBefore"]);
  const OVERLAP = new Set(["Inside", "InsideStart", "InsideEnd", "Equal", "Contains", "ContainsStart",
    "ContainsEnd", "OverlapsBefore", "OverlapsAfter"]);

  let state = null;
  let allPlaces = false; // one-shot: the next Adopted/Failed covers every place
  let onStatus = () => {};
  let onState = () => {};
  let onArm = () => {};
  let onChange = () => {};
  let chain = Promise.resolve();

  // ------------------------------------------------------------------ state

  function loadState() {
    state = Office.context.document.settings.get(STATE_KEY) ||
      { version: 1, seq: 0, activeLevel: 2, openThreadId: null, threads: {} };
  }

  function saveState() {
    Office.context.document.settings.set(STATE_KEY, state);
    Office.context.document.settings.saveAsync();
    onState(state);
  }

  function pendingThreads() {
    return Object.values(state.threads).filter((t) => t.status === "pending");
  }

  function openThread(level) {
    const t = state.threads[state.openThreadId];
    return t && t.status === "pending" && t.level === level ? t : null;
  }

  function draftThread(level, parent) {
    return { id: String(state.seq + 1), seq: state.seq + 1, level, parentId: parent ? parent.id : null, status: "pending", anchors: [] };
  }

  function commitThread(t) {
    if (state.threads[t.id]) return;
    state.seq = t.seq;
    state.threads[t.id] = t;
    state.openThreadId = t.id;
  }

  function newTag(t) {
    return `${TAG_PREFIX}${t.id}-${t.anchors.length + 1}-${Date.now().toString(36)}`;
  }

  // Tags of other pending threads at this level or deeper: their spans are
  // left alone when transforming text that happens to contain them.
  function skipTags(level, except) {
    const tags = [];
    for (const t of pendingThreads()) {
      if (t !== except && t.level >= level) for (const a of t.anchors) tags.push(a.tag);
    }
    return tags;
  }

  function mostRecent(items) {
    return items.reduce((best, x) => (!best || x.thread.seq > best.thread.seq ? x : best), null);
  }

  // Spans recorded before places existed each count as their own place.
  function placeOf(anchor) {
    return anchor.place || anchor.tag;
  }

  function placeCount(anchorList) {
    return new Set(anchorList.map(placeOf)).size;
  }

  // Which place a new span belongs to. It joins a place of its own amendment
  // that it touches or follows with only spaces and punctuation between (a
  // strike and the replacement words typed after it); otherwise it starts a
  // new place. It
  // also links to the place of the parent amendment it touches or sits next
  // to, which is what the order guard follows.
  async function placeFor(context, sel, anchors, thread, parent, tag) {
    const near = await nearSelection(context, sel,
      anchors.filter((a) => a.thread === thread || (parent && a.thread === parent)));
    const sibling = mostRecentAnchor(near.filter((a) => a.thread === thread));
    if (sibling) return { place: placeOf(sibling.anchor), parentPlace: sibling.anchor.parentPlace || null };
    const host = mostRecentAnchor(near.filter((a) => parent && a.thread === parent));
    return { place: tag, parentPlace: host ? placeOf(host.anchor) : null };
  }

  function mostRecentAnchor(items) {
    // Later spans of the same thread were pushed later; prefer the last one.
    return items.length ? items[items.length - 1] : null;
  }

  // Spaces and punctuation, but no words or numbers: what may sit between a
  // strike and its replacement words, such as a comma the strike left.
  const ONLY_SPACES_PUNCT = /^[\s\u00a0\u200b\p{P}]*$/u;

  // Anchors that touch the selection, or are in the same paragraph with only
  // spaces and punctuation between them and the selection.
  async function nearSelection(context, sel, cands) {
    const touching = cands.filter((a) => !["Before", "After", "Unrelated"].includes(a.rel));
    const apart = cands.filter((a) => a.rel === "Before" || a.rel === "After");
    if (!apart.length) return touching;

    // Paragraph has no compareLocationWith; compare the paragraphs' ranges.
    const selPara = sel.paragraphs.getFirst().getRange("Whole");
    const sameParas = apart.map((a) => a.range.paragraphs.getFirst().getRange("Whole").compareLocationWith(selPara));
    await context.sync();
    const sameParaApart = apart.filter((a, i) => sameParas[i].value === "Equal");
    if (!sameParaApart.length) return touching;

    const gaps = sameParaApart.map((a) => {
      const gap = a.rel === "After"
        ? a.range.getRange("End").expandTo(sel.getRange("Start"))
        : sel.getRange("End").expandTo(a.range.getRange("Start"));
      gap.load("text");
      return gap;
    });
    await context.sync();
    return touching.concat(sameParaApart.filter((a, i) => ONLY_SPACES_PUNCT.test(gaps[i].text)));
  }

  // --------------------------------------------------------------- snapshot

  // One look at the document: the selection, and every pending anchor with
  // its position relative to the selection.
  async function snapshot(context) {
    const sel = context.document.getSelection();
    sel.load("isEmpty");
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    const byTag = new Map();
    for (const cc of ccs.items) if (cc.tag && cc.tag.startsWith(TAG_PREFIX)) byTag.set(cc.tag, cc);
    reviveUndoneVotes(byTag);

    const anchors = [];
    for (const thread of pendingThreads()) {
      for (const anchor of thread.anchors) {
        const cc = byTag.get(anchor.tag);
        if (!cc) continue;
        const range = cc.getRange("Content");
        anchors.push({ thread, anchor, cc, range, rel: sel.compareLocationWith(range) });
      }
    }
    await context.sync();
    for (const a of anchors) a.rel = a.rel.value;
    return { sel, anchors };
  }

  // A vote removes the content control of every span it resolves. Ctrl+Z on
  // the vote brings those controls (and their tags) back, so a recorded span
  // whose tag is in the document again was undone: it is pending again, and
  // anything the vote handed over to the parent amendment goes back too.
  // Works whichever votes are undone and in whatever order, and for a partial
  // undo picks up the rest at the next key press.
  function reviveUndoneVotes(byTag) {
    let changed = false;
    for (const vote of state.votes || []) {
      const back = vote.entries.filter((anchor) => byTag.has(anchor.tag));
      const t = state.threads[vote.threadId];
      if (back.length && t) {
        for (const anchor of back) {
          if (!t.anchors.some((a) => a.tag === anchor.tag)) t.anchors.push(anchor);
        }
        const backTags = new Set(back.map((a) => a.tag));
        t.resolved = (t.resolved || []).filter((r) => !backTags.has(r.tag));
        t.status = "pending";
        vote.entries = vote.entries.filter((a) => !backTags.has(a.tag));
        changed = true;
      }
      // A handed-over span whose control is gone was taken out by the undo.
      vote.handedOver = vote.handedOver.filter((h) => {
        if (byTag.has(h.tag) || vote.entries.length) return true;
        const parent = state.threads[h.threadId];
        if (parent) parent.anchors = parent.anchors.filter((a) => a.tag !== h.tag);
        changed = true;
        return false;
      });
    }
    // Spans the empty-span cleanup removed, brought back by Ctrl+Z.
    state.dropped = (state.dropped || []).filter((d) => {
      if (!byTag.has(d.anchor.tag)) return true;
      const t = state.threads[d.threadId];
      if (t) {
        if (!t.anchors.some((a) => a.tag === d.anchor.tag)) t.anchors.push(d.anchor);
        t.status = "pending";
      }
      changed = true;
      return false;
    });
    if (!changed) return;
    state.votes = state.votes.filter((v) => v.entries.length || v.handedOver.length);
    saveState();
  }

  // An insertion span emptied by a later edit (its only words struck by an
  // adopted Level 3 amendment, or struck by its own amendment) would show
  // Word's "Click or tap here to enter text." placeholder. Remove such spans;
  // a thread left with no spans has nothing to vote on and is withdrawn.
  // Removals are remembered (state.dropped) so Ctrl+Z can bring them back.
  // Empty add-in spans that no pending amendment tracks are removed too, so
  // the placeholder never stays on screen.
  async function dropEmptyAnchors(context) {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();
    const byTag = new Map();
    for (const cc of ccs.items) if (cc.tag && cc.tag.startsWith(TAG_PREFIX)) byTag.set(cc.tag, cc);

    const checks = [];
    const tracked = new Set();
    for (const thread of pendingThreads()) {
      for (const anchor of thread.anchors) {
        tracked.add(anchor.tag);
        const cc = byTag.get(anchor.tag);
        if (!cc || anchor.prior) continue;
        // Range.isEmpty is false while Word shows the placeholder, so read
        // the control's OOXML instead.
        checks.push({ thread, anchor, cc, ox: cc.getOoxml() });
      }
    }
    for (const [tag, cc] of byTag) {
      if (!tracked.has(tag)) checks.push({ thread: null, anchor: { tag }, cc, ox: cc.getOoxml() });
    }
    if (!checks.length) return;
    await context.sync();

    const empty = checks.filter((c) => AmdOoxml.isEmptySpan(c.ox.value, c.anchor.tag));
    if (!empty.length) return;
    for (const c of empty) {
      c.cc.delete(false);
      if (!c.thread) continue;
      c.thread.anchors = c.thread.anchors.filter((a) => a !== c.anchor);
      state.dropped = (state.dropped || []).concat({ threadId: c.thread.id, anchor: c.anchor }).slice(-MAX_VOTES);
    }
    await context.sync();
    for (const t of pendingThreads()) {
      if (t.anchors.length) continue;
      t.status = t.resolved && t.resolved.length ? overallOutcome(t) : "withdrawn";
      if (state.openThreadId === t.id) state.openThreadId = null;
    }
  }

  // Word shows "Click or tap here to enter text." in an empty content control,
  // and a span can be empty for a moment while Ctrl+Z undoes a vote step by
  // step. An invisible placeholder (a zero-width space) keeps that text from
  // flashing up. Done in its own sync so a Word build that refuses it can't
  // stop the action itself.
  async function hidePlaceholders(context, ccs) {
    if (!ccs.length) return;
    try {
      for (const cc of ccs) cc.placeholderText = "​";
      await context.sync();
    } catch (err) {
      console.warn("Could not set placeholder text", err);
    }
  }

  function overallOutcome(t) {
    const outcomes = new Set(t.resolved.map((r) => r.outcome));
    return outcomes.size === 1 ? [...outcomes][0] : "mixed";
  }

  // Diagnostic: when Adopted/Failed finds no amendment at the cursor, keep a
  // note of what was checked in the document's add-in settings, where
  // test/dump-docx.ps1 can read it from a saved copy. Invisible in Word.
  function recordMiss(sel, anchors) {
    const checked = anchors.map((a) => `L${a.thread.level} t${a.thread.id} ${a.anchor.tag}: ${a.rel}`);
    const missing = pendingThreads().flatMap((t) => t.anchors)
      .filter((an) => !anchors.some((a) => a.anchor === an)).map((an) => an.tag);
    Office.context.document.settings.set("amendmentLayersLastMiss", {
      at: new Date().toISOString(),
      selectionEmpty: sel.isEmpty,
      checked,
      notInDocument: missing,
    });
    Office.context.document.settings.saveAsync();
  }

  function atCursor(a, selEmpty) {
    return INSIDE.has(a.rel) || (selEmpty && (AT_END.has(a.rel) || AT_START.has(a.rel)));
  }

  function findParent(level, anchors, selEmpty) {
    if (level === 1) return null;
    const near = mostRecent(anchors.filter((a) => a.thread.level === level - 1 && atCursor(a, selEmpty)));
    if (near) return near.thread;
    const any = mostRecent(pendingThreads().filter((t) => t.level === level - 1).map((thread) => ({ thread })));
    return any ? any.thread : null;
  }

  // ---------------------------------------------------------------- marking

  // fromPane: the pane's Insert button, which can't hand the keyboard back to
  // the document, so it only marks words already typed and selected.
  async function mark(kind, fromPane) {
    await Word.run(async (context) => {
      const level = state.activeLevel;
      const { sel, anchors } = await snapshot(context);
      if (kind === "del" && sel.isEmpty) {
        return status("Select the words to strike first, then click Strike.", "warn");
      }

      let thread = openThread(level);
      if (!thread) {
        const parent = findParent(level, anchors, sel.isEmpty);
        if (level === 3 && !parent) {
          return status("There is no pending Level 2 (primary) amendment for this secondary amendment to amend.", "warn");
        }
        thread = draftThread(level, parent);
      }

      if (sel.isEmpty && fromPane) {
        return status("Type the words first, select them, then click Insert.", "warn");
      }
      if (sel.isEmpty) return insertAtCursor(context, anchors, level, thread);

      const skip = skipTags(level, thread);
      const transform = (xml, ownTag) => kind === "del"
        ? AmdOoxml.markDelete(xml, level, skip, ownTag)
        : AmdOoxml.markInsert(xml, level, skip, ownTag);

      // Marking inside a span this thread already owns: transform in place.
      if (anchors.some((a) => a.thread === thread && INSIDE.has(a.rel))) {
        const ox = sel.getOoxml();
        await context.sync();
        const out = transform(ox.value, null);
        if (out.empty) sel.delete();
        else sel.insertOoxml(out.xml, "Replace").select("End");
        await context.sync();
        await dropEmptyAnchors(context);
        saveState();
        return status(`${kind === "del" ? "Struck" : "Marked inserted"} at ${LEVEL_NAMES[level]}.`);
      }

      const tag = newTag(thread);
      const parent = thread.parentId ? state.threads[thread.parentId] : null;
      const where = await placeFor(context, sel, anchors, thread, parent, tag);
      const ox = sel.getOoxml();
      const cc = sel.insertContentControl();
      cc.tag = tag;
      cc.title = LEVEL_NAMES[level];
      cc.appearance = "Hidden";
      await context.sync();
      await hidePlaceholders(context, [cc]);

      const out = transform(ox.value, tag);
      if (out.empty) {
        // Striking only this level's own insertion takes it back out; there
        // is no span left to record.
        cc.delete(false);
        await context.sync();
        await dropEmptyAnchors(context);
        saveState();
        return status(`Removed at ${LEVEL_NAMES[level]}.`);
      }
      cc.insertOoxml(out.xml, "Replace");
      cc.getRange("Content").select("End");
      await context.sync();

      // Text marked as an insertion is the amendment's own new wording (typed,
      // then selected and marked), so failing the amendment removes it.
      thread.anchors.push({ tag, prior: kind === "ins" ? "" : AmdOoxml.slim(ox.value), ...where });
      commitThread(thread);
      await dropEmptyAnchors(context);
      saveState();
      status(`${kind === "del" ? "Struck" : "Marked inserted"} at ${LEVEL_NAMES[level]}.`);
    });
  }

  // Insert key with nothing selected: put a one-space placeholder in the
  // level's insertion style at the cursor and select it, so the staff
  // member's typing replaces it and inherits the formatting.
  async function insertAtCursor(context, anchors, level, thread) {
    const own = mostRecent(anchors.filter((a) => a.thread === thread && atCursor(a, true)));
    // Otherwise go inside the deepest lower-level span the cursor touches, so
    // the new words sit within the text they amend.
    const lower = anchors.filter((a) => a.thread.level < level && atCursor(a, true));
    const deepest = Math.max(0, ...lower.map((a) => a.thread.level));
    const host = own || mostRecent(lower.filter((a) => a.thread.level === deepest));
    const sel = context.document.getSelection();
    const tag = newTag(thread);
    const parent = thread.parentId ? state.threads[thread.parentId] : null;
    const placement = own ? null : await placeFor(context, sel, anchors, thread, parent, tag);

    let target = sel;
    let where = "Start";
    if (host && AT_END.has(host.rel)) { target = host.cc; where = "End"; }
    else if (host && AT_START.has(host.rel)) { target = host.cc; where = "Start"; }

    // Continuing this thread's own span at its end: keep a plain space
    // between the struck text and the new words.
    const withSep = !!own && where === "End";
    const sepRange = withSep ? target.insertText(" ", where) : null;
    const phRange = sepRange ? sepRange.insertText(" ", "After") : target.insertText(" ", where);
    const sepOx = sepRange ? sepRange.getOoxml() : null;
    const phOx = phRange.getOoxml();
    await context.sync();

    if (sepRange) sepRange.insertOoxml(AmdOoxml.restyleAll(sepOx.value, "plain", 0, true).xml, "Replace");
    const placed = phRange.insertOoxml(AmdOoxml.restyleAll(phOx.value, "ins", level, false).xml, "Replace");

    if (own) {
      placed.select();
      await context.sync();
    } else {
      const cc = placed.insertContentControl();
      cc.tag = tag;
      cc.title = LEVEL_NAMES[level];
      cc.appearance = "Hidden";
      cc.getRange("Content").select();
      await context.sync();
      await hidePlaceholders(context, [cc]);
      thread.anchors.push({ tag, prior: "", ...placement });
      commitThread(thread);
      saveState();
    }
    status(`Type the ${LEVEL_NAMES[level]} insertion.`);
  }

  // -------------------------------------------------------------- resolving

  async function resolve(adopted) {
    await Word.run(async (context) => {
      const { sel, anchors } = await snapshot(context);
      const hit = mostRecent(anchors.filter((a) => atCursor(a, sel.isEmpty)));
      if (!hit) {
        recordMiss(sel, anchors);
        if (state.activeLevel === 1) return resolveUntracked(context, sel, anchors, adopted);
        const latest = mostRecent(pendingThreads().map((thread) => ({ thread })));
        if (!latest) return status("No amendment is pending.", "warn");
        return status("Cursor is not in the immediately pending amendment. Click inside the text of the " +
          `pending ${LEVEL_NAMES[latest.thread.level]} first.`, "warn");
      }
      const t = hit.thread;
      const every = allPlaces;
      const here = placeOf(hit.anchor);
      const inScope = (anchor) => every || placeOf(anchor) === here;
      const places = new Set(t.anchors.filter(inScope).map(placeOf));

      // Order guard: an amendment to this one, still pending at one of the
      // places being resolved. Pending amendments elsewhere don't block, nor
      // do marks no longer in the document (taken out with Ctrl+Z).
      const child = mostRecent(anchors.filter((a) => a.thread.parentId === t.id &&
        (a.anchor.parentPlace == null || places.has(a.anchor.parentPlace))));
      if (child) {
        return status("Cursor in wrong amendment level: the " + `${LEVEL_NAMES[child.thread.level]} to this ` +
          `amendment is still pending${every ? " in at least one place" : ""}. ` +
          "Click inside its text and resolve it first.", "error");
      }

      const mine = anchors.filter((a) => a.thread === t && inScope(a.anchor));
      const others = anchors.filter((a) => a.thread !== t && a.thread.level >= t.level);
      const parent = t.parentId ? state.threads[t.parentId] : null;
      const parentAnchors = parent && parent.status === "pending" ? anchors.filter((a) => a.thread === parent) : [];

      const nested = [];
      for (const m of mine) {
        for (const o of others) nested.push(o.range.compareLocationWith(m.range));
        m.inParent = parentAnchors.map((pa) => m.range.compareLocationWith(pa.range));
        if (adopted) m.ox = m.range.getOoxml();
      }
      await context.sync();

      if (nested.some((r) => INSIDE.has(r.value))) {
        return status("Another pending amendment is marked inside this one. Resolve it first.", "error");
      }

      // Every span this vote resolves loses its content control. If Ctrl+Z
      // later brings one back, reviveUndoneVotes treats it as pending again.
      const outcome = adopted ? "adopted" : "failed";
      const vote = { id: (state.voteSeq = (state.voteSeq || 0) + 1), threadId: t.id, entries: [], handedOver: [] };
      const holders = [];

      if (adopted) {
        const skip = skipTags(t.level, t);
        const outs = mine.map((m) => AmdOoxml.fold(m.ox.value, t.level, skip, m.anchor.tag, false));
        const problems = outs.flatMap((o) => o.problems);
        if (problems.length) return status(`Nothing changed: ${problems[0]}`, "error");

        mine.forEach((m, i) => {
          const out = outs[i];
          if (out.empty) return m.cc.delete(false);
          const folded = m.cc.insertOoxml(out.xml, "Replace");
          m.cc.delete(true);
          // Text this amendment added outside the parent's spans now belongs
          // to the parent, so the parent's own vote can fold or revert it. It
          // gets a new content control, so Ctrl+Z on this vote takes the
          // handover back out along with everything else.
          if (parent && parentAnchors.length && !m.inParent.some((r) => INSIDE.has(r.value))) {
            const parentPlace = m.anchor.parentPlace || null;
            const peer = parent.anchors.find((an) => parentPlace && placeOf(an) === parentPlace);
            const tag = newTag(parent);
            const holder = folded.insertContentControl();
            holder.tag = tag;
            holder.title = LEVEL_NAMES[parent.level];
            holder.appearance = "Hidden";
            holders.push(holder);
            parent.anchors.push({
              tag,
              prior: m.anchor.prior,
              place: parentPlace || tag,
              parentPlace: peer ? peer.parentPlace || null : null,
            });
            vote.handedOver.push({ threadId: parent.id, tag });
          }
        });
      } else {
        for (const m of mine) {
          if (!m.anchor.prior) {
            m.cc.delete(false);
          } else {
            m.cc.insertOoxml(AmdOoxml.forInsert(m.anchor.prior, m.anchor.tag), "Replace");
            m.cc.delete(true);
          }
        }
      }
      await context.sync();
      await hidePlaceholders(context, holders);

      // Spans of these places that are no longer in the document are dropped
      // along with the ones just resolved.
      const done = t.anchors.filter(inScope);
      const missing = done.length - mine.length;
      t.anchors = t.anchors.filter((a) => !inScope(a));
      t.resolved = (t.resolved || []).concat(done.map((a) => ({ tag: a.tag, outcome })));
      vote.entries = mine.map((m) => m.anchor);
      state.votes = (state.votes || []).concat(vote).slice(-MAX_VOTES);
      if (!t.anchors.length) t.status = overallOutcome(t);
      state.openThreadId = null; // a vote ends marking; the next mark is a new amendment
      if (every) setAllPlaces(false);
      await dropEmptyAnchors(context);
      saveState();

      const left = placeCount(anchors.filter((a) => a.thread === t && !inScope(a.anchor)).map((a) => a.anchor));
      status(`${LEVEL_NAMES[t.level]} ${outcome}` +
        (every && places.size > 1 ? ` in all ${places.size} places.` : left ? " here." : ".") +
        (left ? ` ${left} other place${left === 1 ? "" : "s"} still pending.` : "") +
        (missing > 0 ? ` (${missing} marked span(s) were no longer in the document.)` : ""));
    });
  }

  // Level 1 markup that was in the document before the meeting has no stored
  // prior text; it is read from its formatting instead. Scope: the one change
  // at the cursor (struck and inserted words with only spaces between), or
  // everything selected.
  async function resolveUntracked(context, sel, anchors, adopted) {
    if (sel.isEmpty) return resolveUntrackedAtCursor(context, sel, adopted);
    const range = sel;
    const checks = anchors.map((a) => a.range.compareLocationWith(range));
    const ox = range.getOoxml();
    await context.sync();

    const blocking = mostRecent(anchors.filter((a, i) => OVERLAP.has(checks[i].value)));
    if (blocking) {
      return status("Cursor in wrong amendment level: the " + `${LEVEL_NAMES[blocking.thread.level]} to this ` +
        "passage is still pending. Click inside its text and resolve it first.", "error");
    }
    const out = adopted ? AmdOoxml.fold(ox.value, 1, [], null, true) : AmdOoxml.failUntracked(ox.value, []);
    if (out.problems && out.problems.length) return status(`Nothing changed: ${out.problems[0]}`, "error");
    if (!out.changed) return status("No Level 1 markup in this passage.", "warn");
    if (out.empty) range.delete();
    else range.insertOoxml(out.xml, "Replace");
    await context.sync();
    status(`Level 1 change ${adopted ? "adopted" : "failed"} in this selection.`);
  }

  async function resolveUntrackedAtCursor(context, sel, adopted) {
    const para = sel.paragraphs.getFirst();
    const range = para.getRange("Content");
    const before = para.getRange("Start").expandTo(sel.getRange("Start"));
    before.load("text");
    const ox = range.getOoxml();
    await context.sync();

    const pending = pendingThreads().flatMap((t) => t.anchors.map((a) => a.tag));
    const out = AmdOoxml.resolveUntrackedAt(ox.value, before.text.length, adopted, pending);
    if (out.blocked) {
      const owner = pendingThreads().find((t) => t.anchors.some((a) => a.tag === out.blocked));
      return status("Cursor in wrong amendment level: the " +
        `${owner ? LEVEL_NAMES[owner.level] : "amendment"} to this change is still pending. ` +
        "Click inside its text and resolve it first.", "error");
    }
    if (!out.changed) return status("No Level 1 change at the cursor. Click inside its struck or inserted words.", "warn");
    if (out.empty) range.delete();
    else range.insertOoxml(out.xml, "Replace");
    await context.sync();
    status(`Level 1 change ${adopted ? "adopted" : "failed"}.`);
  }

  // ---------------------------------------------------------------- plumbing

  function status(text, kind) {
    onStatus(text, kind || "ok");
  }

  function setAllPlaces(on) {
    allPlaces = on;
    onArm(on);
  }

  function describeError(err) {
    const msg = (err && err.message) || String(err);
    if (/content control|InvalidArgument|GeneralException/i.test(msg)) {
      return "Word couldn't mark that selection. If it partly overlaps another amendment's text, select only inside it or only outside it.";
    }
    return `Something went wrong: ${msg}`;
  }

  // Keystrokes can arrive faster than Word finishes; run actions one at a time.
  // After each one, onChange lets Presenter view refresh.
  function queue(fn) {
    chain = chain.then(fn).catch((err) => {
      console.error(err);
      status(describeError(err), "error");
    }).then(() => onChange());
    return chain;
  }

  // Runs other Word work (Presenter view's reads) in the same line, so it
  // never interleaves with an action. Errors go back to the caller.
  function serial(fn) {
    const result = chain.then(fn);
    chain = result.catch(() => {});
    return result;
  }

  return {
    LEVEL_NAMES,
    init(handlers) {
      onStatus = handlers.onStatus || onStatus;
      onState = handlers.onState || onState;
      onArm = handlers.onArm || onArm;
      onChange = handlers.onChange || onChange;
      loadState();
      onState(state);
    },
    get activeLevel() {
      return state.activeLevel;
    },
    setLevel(level) {
      state.activeLevel = level;
      state.openThreadId = null; // the next mark starts a new amendment
      setAllPlaces(false);
      saveState();
      status(`${LEVEL_NAMES[level]} is active. The next mark starts a new amendment.`);
    },
    toggleAllPlaces() {
      setAllPlaces(!allPlaces);
      status(allPlaces
        ? "All places: the next Adopted or Failed applies to every place of the amendment at the cursor."
        : "All places is off. Adopted and Failed apply to the place at the cursor.");
    },
    serial,
    markDelete: () => queue(() => mark("del")),
    markInsert: () => queue(() => mark("ins")),
    markInsertFromPane: () => queue(() => mark("ins", true)),
    adopted: () => queue(() => resolve(true)),
    failed: () => queue(() => resolve(false)),
  };
})();
