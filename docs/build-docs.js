// Builds the staff documents in this folder:
//   Amendment Layers - Quick Reference.docx
//   Amendment Layers - User Manual.docx
// Run from the project folder:  node docs/build-docs.js
// The manual's Contents is a Word table of contents (page numbers and links),
// built from its Heading 1 headings. After writing the manual, this script has
// Word fill it in (docs/update-toc.ps1), so Word must be installed.
//
// The documents assume the staff member uses the pane's buttons; the
// keyboard shortcuts are an alternative. American spelling throughout.
//
// Inline markup in the text below:
//   **bold**   *italic*   `Ctrl+Alt+Shift+D` (a key)
//   [[ins2:words]] / [[del3:words]]   sample text in a level's formatting
//   [^3]   footnote reference (the manual's footnotes are in MANUAL_FOOTNOTES)
//   {{warn:words}}   a warning: bold italic red
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const {
  AlignmentType, BorderStyle, Document, Footer, FootnoteReferenceRun, LevelFormat, Packer, PageNumber, Paragraph,
  ShadingType, Table, TableCell, TableOfContents, TableRow, TextRun, UnderlineType, VerticalMergeType, WidthType,
  HeadingLevel,
} = require("docx");

const FONT = "Calibri";
const WARN_RED = "C00000";
const SHADE = { 2: "DCE9FA", 3: "F5DDF5" };
// All text black; the level's color is only in the shading and underline.
// del2u1 / del3u2: words a lower level inserted, struck where they stand and
// keeping that level's underline.
const UNDER = {
  1: { type: UnderlineType.SINGLE, color: "000000" },
  2: { type: UnderlineType.DOUBLE, color: "1F5FBF" },
  3: { type: UnderlineType.DOTTED, color: "8B1E8B" },
};
const LEVEL_RUN = {
  ins1: { underline: UNDER[1], color: "000000" },
  del1: { strike: true, color: "000000" },
  ins2: { underline: UNDER[2], color: "000000", shading: shade(2) },
  del2: { doubleStrike: true, color: "000000", shading: shade(2) },
  del2u1: { doubleStrike: true, underline: UNDER[1], color: "000000", shading: shade(2) },
  ins3: { italics: true, underline: UNDER[3], color: "000000", shading: shade(3) },
  del3: { italics: true, strike: true, color: "000000", shading: shade(3) },
  del3u2: { italics: true, strike: true, underline: UNDER[2], color: "000000", shading: shade(3) },
};

const SITE = "https://amendment-layers.netlify.app";
const PROJECT = "C:\\Users\\burke\\Code by Claude\\amendment-display-addin-v4";

function shade(level) {
  return { type: ShadingType.CLEAR, color: "auto", fill: SHADE[level] };
}

// ------------------------------------------------------------------ helpers

function runs(text, base = {}) {
  const out = [];
  const re = /\*\*(.+?)\*\*|\*(.+?)\*|\[\[(ins|del)([123](?:u[12])?):(.+?)\]\]|`(.+?)`|\[\^(\d+)\]|\{\{warn:(.+?)\}\}/g;
  let last = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ ...base, text: text.slice(last, m.index) }));
    if (m[8]) out.push(new TextRun({ ...base, text: m[8], bold: true, italics: true, color: WARN_RED }));
    else if (m[7]) out.push(new FootnoteReferenceRun(Number(m[7])));
    else if (m[1]) out.push(new TextRun({ ...base, text: m[1], bold: true }));
    else if (m[2]) out.push(new TextRun({ ...base, text: m[2], italics: true }));
    else if (m[3]) out.push(new TextRun({ ...base, text: m[5], ...LEVEL_RUN[m[3] + m[4]] }));
    else out.push(new TextRun({ ...base, text: m[6], bold: true, font: "Consolas" }));
    last = re.lastIndex;
  }
  if (last < text.length) out.push(new TextRun({ ...base, text: text.slice(last) }));
  return out;
}

function makeKit(size) {
  // size: body text in half-points
  const base = { font: FONT, size };
  let listInstance = 0;

  const p = (text, opts = {}) => new Paragraph({
    children: runs(text, { ...base, ...(opts.run || {}) }),
    spacing: { after: opts.after ?? 100, before: opts.before ?? 0 },
    keepNext: opts.keepNext,
    alignment: opts.align,
  });

  const heading = (level, text) => new Paragraph({
    heading: level,
    children: [new TextRun({ text, font: FONT })],
    keepNext: true,
  });

  const bullets = (items, ref = "bullets") => items.map((t) => new Paragraph({
    children: runs(t, base),
    numbering: { reference: ref, level: 0 },
    spacing: { after: 60 },
  }));

  const steps = (items) => {
    listInstance++;
    return items.map((t) => new Paragraph({
      children: runs(t, base),
      numbering: { reference: "steps", level: 0, instance: listInstance },
      spacing: { after: 60 },
    }));
  };

  // opts: fill, bold, span (columns), merge ("restart" | "continue" down rows)
  const cell = (content, width, opts = {}) => new TableCell({
    width: { size: width, type: WidthType.DXA },
    columnSpan: opts.span,
    verticalMerge: opts.merge === "restart" ? VerticalMergeType.RESTART
      : opts.merge === "continue" ? VerticalMergeType.CONTINUE : undefined,
    shading: opts.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opts.fill } : undefined,
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    children: (Array.isArray(content) ? content : [content]).map((t) => new Paragraph({
      children: runs(t, { ...base, bold: opts.bold, size: opts.size || base.size }),
      spacing: { after: 30 },
    })),
  });

  // rows: arrays of cell contents; first row is the header.
  const table = (widths, rows) => new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, i) => new TableRow({
      tableHeader: i === 0,
      cantSplit: true,
      children: row.map((c, j) => cell(c, widths[j], i === 0 ? { fill: "E7E6E6", bold: true } : {})),
    })),
  });

  // The legend, with Level 1's formatting cells split: below them, merged
  // across both, the note on when Level 1 markup appears.
  const legend = (widths) => {
    const [w1, w2, w3, w4] = widths;
    const row = (cells) => new TableRow({ cantSplit: true, children: cells });
    return new Table({
      width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
      columnWidths: widths,
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true,
          children: LEGEND_HEAD.map((t, j) => cell(t, widths[j], { fill: "E7E6E6", bold: true })) }),
        row([cell("**1**", w1, { merge: "restart" }), cell("Main motion", w2, { merge: "restart" }),
          cell("[[ins1:black, underlined]]", w3), cell("[[del1:black, struck through]]", w4)]),
        row([cell("", w1, { merge: "continue" }), cell("", w2, { merge: "continue" }),
          cell(`*${LEVEL1_NOTE}*`, w3 + w4, { span: 2, size: Math.max(16, base.size - 2) })]),
        row([cell("**2**", w1), cell("Primary amendment", w2),
          cell("[[ins2:blue double underline, blue shading]]", w3), cell("[[del2:double strike, blue shading]]", w4)]),
        row([cell("**3**", w1), cell("Secondary amendment", w2),
          cell("[[ins3:italic, purple dotted underline, purple shading]]", w3), cell("[[del3:italic, struck through, purple shading]]", w4)]),
      ],
    });
  };

  // A boxed note (single-cell table with a heavy left rule).
  const callout = (width, title, lines, fill = "FFF4C2") => new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths: [width],
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        width: { size: width, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: "auto", fill },
        margins: { top: 90, bottom: 90, left: 160, right: 160 },
        borders: {
          left: { style: BorderStyle.SINGLE, size: 24, color: "B7791F" },
          top: { style: BorderStyle.SINGLE, size: 4, color: "B7791F" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "B7791F" },
          right: { style: BorderStyle.SINGLE, size: 4, color: "B7791F" },
        },
        children: [
          new Paragraph({ children: runs(`**${title}**`, base), spacing: { after: 50 } }),
          ...lines.map((t) => new Paragraph({ children: runs(t, base), spacing: { after: 50 } })),
        ],
      })],
    })],
  });

  const gap = (after = 120) => new Paragraph({ children: [], spacing: { after } });

  return { p, heading, bullets, steps, table, legend, callout, gap, base };
}

function numbering() {
  return {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 240 } } } }],
      },
      {
        reference: "steps",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 300 } } } }],
      },
    ],
  };
}

function styles(bodySize, h1, h2) {
  return {
    default: { document: { run: { font: FONT, size: bodySize } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", next: "Normal",
        run: { font: FONT, size: h1 + 12, bold: true, color: "1B1B1B" }, paragraph: { spacing: { after: 60 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: h1, bold: true, color: "1F3864" },
        paragraph: { spacing: { before: 200, after: 90 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: h2, bold: true, color: "1F3864" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      // Contents entries (Word's "TOC 1").
      { id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal",
        run: { font: FONT, size: bodySize }, paragraph: { spacing: { after: 20 } } },
    ],
  };
}

const LETTER = { width: 12240, height: 15840 };

function footer(text) {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ font: FONT, size: 16, color: "777777", children: [`${text}  \u00b7  Page `, PageNumber.CURRENT] })],
    })],
  });
}

// ------------------------------------------------------------- shared text

// strikeNote: footnote reference after "selected" (the manual only).
const buttons = (strikeNote = "") => [
  ["Pane button", "What it does", "Keyboard shortcuts"],
  ["**1**, **2**, **3**", "Set the level for a new amendment (1 main motion, 2 primary, 3 secondary)", "\u2014"],
  ["**Strike**", `Strike the selected${strikeNote} words`, "`Ctrl+Alt+Shift+D`"],
  ["**Insert**", "Mark the selected words as inserted", "`Ctrl+Alt+Shift+I`"],
  ["**Adopted**", "Record \u201cadopted\u201d for the amendment at the cursor", "`Ctrl+Alt+Shift+Y`"],
  ["**Failed**", "Record \u201cfailed\u201d for the amendment at the cursor", "`Ctrl+Alt+Shift+N`"],
  ["**All places**", "Make the next Adopted/Failed apply in every place", "\u2014"],
  ["**Presenter view**", "Open or close the window for the projected screen", "\u2014"],
  ["**Tracked Changes \u2192 manual markup**", "Turn Tracked Changes into struck/underlined text", "\u2014"],
  ["**Manual markup \u2192 Tracked Changes**", "Turn struck/underlined text into Tracked Changes, after all votes", "\u2014"],
];

const LEGEND_HEAD = ["Level", "Represents", "Inserted words", "Struck words"];
const LEVEL1_NOTE = "(only if the main motion is one to Rescind or Amend Something Previously Adopted)";

const SEVERAL_PLACES = [
  "When one motion makes a change in several places (for example, \u201cBudget Committee\u201d to \u201cFinance Committee\u201d everywhere it appears), mark **every** place **without clicking the level button in between**.",
  "Clicking a level button starts a **new** amendment. If you click it between places, each place becomes a separate amendment, voted on its own, even if “All places” is clicked.",
];

const RESULTS = [
  ["Level voted on", "Adopted", "Failed"],
  ["**3** Secondary", "Its change is made in the primary amendment\u2019s text: words it struck that the primary amendment proposed to insert disappear; its new words take the primary amendment\u2019s formatting. The purple marks disappear.", "The purple marks (strikethrough, italics, shading) disappear; the primary amendment\u2019s text is as it was."],
  ["**2** Primary", "Its change is made in the main motion\u2019s text: words it struck that the main motion proposed to insert disappear; other words it struck are shown with a strikeout line in the main motion. The blue marks disappear.", "The blue marks (double strikethrough, shading) disappear; the main motion\u2019s text is as it was."],
  ["**1** Main motion", "Struck words are removed; the underline is removed from inserted words. A wholly struck paragraph goes, section number and all.", "Inserted words are removed; struck words have the strikeout line removed from them."],
];

// The manual's footnotes, numbered in the order they appear in the text.
const MANUAL_FOOTNOTES = {
  1: "Henry M. Robert et al., *Robert’s Rules of Order Newly Revised*, 12th edition (Public Affairs, 2020), vii, cited as “RONR (12th ed.).”",
  2: "Henry M. Robert et al., *Robert’s Rules of Order, Newly Revised in Brief*, 3rd edition (Public Affairs, 2020), 19–24, cited as “RONRIB (3rd ed.).” RONRIB is a short, user-friendly work to “gain an introductory familiarity with meeting rules.” RONR (12th ed.) p. xxiv.",
  3: "RONRIB (3rd ed.) 60–62.",
  4: "RONRIB (3rd ed.) 53–55.",
  5: "The most common way to “select” text in Microsoft Word is to click at the start of the text, hold down the left mouse button, and drag the cursor over the text. For additional or alternate ways, consult Microsoft Help.",
  6: "This procedure applies not only when you need to undo a mistake but also when the members take action to retake or undo a vote. For example, after the presiding officer announces the result of a voice vote, a motion may be adopted to retake the vote as a counted vote, or the members may adopt a motion to *Reconsider* the vote. See RONRIB (3rd ed.) 70–76 (other methods of voting); 58–60 (*Reconsider*).",
};

function footnotes(notes) {
  const out = {};
  for (const [id, text] of Object.entries(notes)) {
    out[id] = { children: [new Paragraph({ children: runs(text, { font: FONT, size: 18 }), spacing: { after: 40 } })] };
  }
  return out;
}

// --------------------------------------------------------- quick reference

function quickReference() {
  const W = 10224; // content width with 0.7" margins
  const k = makeKit(19);
  const children = [
    new Paragraph({ style: "Title", children: [new TextRun({ text: "Amendment Layers \u2014 Quick Reference", font: FONT })] }),
    k.p("Live display of amendments in Word. Keep this card beside the keyboard. Full details: *Amendment Layers \u2014 User Manual*.", { after: 120 }),

    k.heading(HeadingLevel.HEADING_1, "Before amendments are offered"),
    ...k.steps([
      "Open a document containing a main motion in Word, then click **Home \u2192 Amendment Layers** to open the pane. (It reopens by itself in documents that have used it before.)",
      "**If the document has Tracked Changes:** {{warn:make a copy of the document first!}} Then click **Tracked Changes \u2192 manual markup** in the pane (manual, section 8).",
      "**For the projected screen:** click **Presenter view** in the pane. Click the new window\u2019s title bar, then press **Windows key + Shift + \u2192** (or \u2190) to move it to the projected screen.",
    ]),

    k.heading(HeadingLevel.HEADING_1, "The pane\u2019s buttons"),
    k.table([1900, 5524, 2800], buttons()),

    k.heading(HeadingLevel.HEADING_1, "What the screen shows"),
    k.legend([900, 2600, 3362, 3362]),

    k.heading(HeadingLevel.HEADING_1, "When an amendment is offered"),
    ...k.steps([
      "Click its level in the pane: **2** for a primary amendment, **3** for a secondary amendment. Click it **even if it is already lit**: each click starts a new amendment.",
      "**To strike:** select the words in the document, then click **Strike**.",
      "**To insert:** type the new words in the document where they belong, select them, then click **Insert**.",
    ]),
    k.gap(60),
    k.callout(W, "One motion, several places", SEVERAL_PLACES),

    k.heading(HeadingLevel.HEADING_1, "When the chair announces the result"),
    ...k.steps([
      "Click inside the text of the amendment just voted on: its struck words or its new words.",
      "Click **Adopted** or **Failed**. There is no confirmation.",
      "**Marked in several places?** Repeat in each place. *Or* click **All places** first, then Adopted or Failed once: it applies everywhere. All places switches itself off after that one vote.",
    ]),
    k.p("**Order:** a secondary amendment is voted before its primary amendment, and a primary amendment before the main motion. If you click in the wrong one, the pane says so and nothing changes.", { before: 60 }),
    // Without the manual's note on wholly struck paragraphs, to keep to two pages.
    k.table([2200, 4012, 4012], RESULTS.map((row) => row.map((cell) => cell.replace(/ A wholly struck paragraph goes, section number and all\./, "")))),

    k.heading(HeadingLevel.HEADING_1, "Main motion (Level 1)"),
    k.p("Click **1** in the pane, click inside the change\u2019s struck or inserted words, then click **Adopted** or **Failed**. Only that change is resolved; to resolve several together, select them all first."),

    k.heading(HeadingLevel.HEADING_1, "Converting Tracked Changes"),
    k.p("{{warn:Work on a copy of the document!}} Each button acts at once on the selected text, or on the whole document if nothing is selected. **Manual markup \u2192 Tracked Changes** needs all votes recorded before it is clicked and leaves Track Changes on (manual, section 8)."),

    // The status box's own words (all are in the manual, section 10).
    k.heading(HeadingLevel.HEADING_1, "Messages in the status box"),
    k.table([4100, W - 4100], [
      ["Message", "What to do"],
      ["\u201cCursor in wrong amendment level\u2026\u201d", "An amendment to this one is still pending here. Resolve the level it names first."],
      ["\u201cCursor is not in the immediately pending amendment\u2026\u201d", "Click inside the text of the amendment it names, then click the button again."],
      ["\u201cSelect the words to strike first\u2026\u201d / \u201cType the words first\u2026\u201d", "Select the words (for Insert, type them first), then click the button again."],
      ["\u201cWord couldn\u2019t mark that selection\u2026\u201d", "Select within one paragraph, inside or outside other amendments."],
      ["\u201cTrack Changes is on\u2026\u201d / \u201cThis text has Tracked Changes\u2026\u201d", "Click **Tracked Changes \u2192 manual markup** first."],
    ]),

    k.heading(HeadingLevel.HEADING_1, "Troubleshooting"),
    k.table([4100, W - 4100], [
      ["Problem", "What to do"],
      ["Wrong words marked", "Click just to the **right** of the text, press `Ctrl+Z` until the mark is gone, then mark again."],
      ["Wrong vote clicked", "Click just to the **right** of the text, press `Ctrl+Z` until it is as before, then vote again."],
      ["Presenter window doesn\u2019t fit the projected screen", "Drag its edges or corners to fit (manual, section 7)."],
    ]),
    // Word needs a paragraph after a final table; keep it 1 pt high so it
    // can't spill onto a page of its own.
    new Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: "exact" }, children: [new TextRun({ text: "", size: 2 })] }),
  ];

  return new Document({
    creator: "Amendment Layers",
    title: "Amendment Layers \u2014 Quick Reference",
    styles: styles(19, 26, 22),
    numbering: numbering(),
    sections: [{
      properties: { page: { size: LETTER, margin: { top: 720, bottom: 720, left: 1008, right: 1008 } } },
      footers: { default: footer("Amendment Layers \u2014 Quick Reference") },
      children,
    }],
  });
}

// ------------------------------------------------------------------ manual

function manual() {
  const W = 9360; // content width with 1" margins
  const k = makeKit(22);
  const H1 = (t) => k.heading(HeadingLevel.HEADING_1, t);
  const H2 = (t) => k.heading(HeadingLevel.HEADING_2, t);

  const children = [
    new Paragraph({ style: "Title", children: [new TextRun({ text: "Amendment Layers \u2014 User Manual", font: FONT })] }),
    k.p("A Microsoft Word add-in for displaying pending amendments live during committee, board, and members\u2019 meetings.", { after: 200 }),
    k.p("**Contents**", { after: 60 }),
    new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-1" }),
    k.gap(),

    H1("1. What the add-in does"),
    k.p("At a meeting for which the parliamentary authority is *Robert\u2019s Rules*,[^1] a proposal for consideration is normally offered as a \u201cmain motion\u201d; after it is \u201cstated\u201d by the presiding officer, it is open for debate and for changes in it to be proposed as \u201camendments.\u201d[^2] There can be \u201cprimary amendments\u201d that propose changes in the main motion, and while a primary amendment is pending (that is, has been stated and opened for debate) a \u201csecondary amendment\u201d may be proposed to change wording in the primary amendment."),
    k.p("There is also a particular type of main motion called to *Rescind or Amend Something Previously Adopted* which, as its name suggests, may seek to change something that has already been adopted, such as a bylaw amendment that proposes changes in the previously adopted bylaws.[^3]"),
    k.p("Finally, there are other types of motions that can be amended, such as to refer pending motions to a committee or to postpone them.[^4] (For the purpose of projection with the use of this add-in, these motions are treated like \u201cmain motions\u201d even though, technically, parliamentary procedure classifies them differently.)"),
    k.p("Amendments should be proposed in forms such as \u201cI move to strike \u2018in 3 years\u2019,\u201d \u201cI move to insert \u2018calendar\u2019 before \u2018years\u2019,\u201d or \u201cI move to strike \u20183\u2019 and insert \u20185\u2019.\u201d However, it is often easiest for those who will debate and vote on amendments to see them in context: that is, to see how the motion proposed to be amended would be affected if the amendment were adopted."),
    k.p("This add-in is designed to assist in projecting, for the members to see, proposed motions that are subject to amendment, with amendments to them as they are proposed and, through votes, adopted or defeated. The projection may be through a shared screen for an online meeting or, in an in-person meeting, in a manner that is visible to its voting members, either directly to a monitor, or to a digital projector that projects an image onto a blank screen."),
    k.p("During a meeting, the document with the pending \u201cmain motion\u201d is projected while amendments to it are proposed and voted on. The add-in shows each layer of amendment in its own formatting, applied with one click, and shows the effect of each vote with one click:"),
    ...k.bullets([
      "**Level 1 \u2014 the main motion** (as described above).",
      "**Level 2 \u2014 a primary amendment** to the main motion.",
      "**Level 3 \u2014 a secondary amendment** to a primary amendment.",
    ]),
    k.p("When the result of a vote on an amendment is announced, one click records it: **Adopted** makes the amendment\u2019s change part of the text one level down; **Failed** removes the amendment and puts the text back as it was before the amendment was moved. Buttons that designate the type of amendment and record voting results are on an add-in “pane” located in a side panel."),
    k.p("For the members, for easier viewing, **Presenter view** shows the document with any pending amendments by itself (without the add-in pane), in large type, on the projected screen, while you keep Word with both the document and the add-in pane on your own screen (section 7)."),
    k.p("The add-in does not use Microsoft Word\u2019s Track Changes to show amendments. All of its formatting is ordinary text formatting, which is why it can show amendments to amendments. It can, however, convert a document\u2019s Tracked Changes into its own formatting, and convert its formatting into Tracked Changes after the votes (section 8)."),
    k.p("If the add-in is not yet installed on your computer, see section 12 for instructions on how to do that."),

    H1("2. Getting started at the meeting"),
    ...k.steps([
      "Open a document containing a \u201cmain motion\u201d in Word. The computer must be connected to the internet: the add-in\u2019s pane loads from the web each time.",
      "Click **Home \u2192 Amendment Layers**. (If \u201cAmendment Layers\u201d is not on the Home ribbon, click **Add-ins** and \u201cAmendment Layers\u201d should appear.) The pane opens on the right. In a document that has used the add-in before, the pane opens by itself.",
      "If the document contains any Tracked Changes, see section 8.",
      "If the document is to be projected without a visible add-in pane to a secondary monitor which will be used for a shared screen in an online meeting or, during an in-person meeting, linked to a digital projector or directly to a monitor visible to the members, open **Presenter view** (section 7).",
    ]),
    H2("The pane"),
    ...k.bullets([
      "**Active level** \u2014 three buttons: **1 Main motion**, **2 Primary**, **3 Secondary**. The lit one is the level your next mark will use. Clicking a button also starts a new amendment (section 4).",
      "**Status box** \u2014 the gray box under the level buttons. After each action it says what happened, or why nothing happened.",
      "**Amendment Actions** \u2014 the **Strike**, **Insert**, **Adopted** and **Failed** buttons. Each shows its keyboard shortcut (section 13).",
      "**All places** \u2014 a one-shot switch for amendments marked in several places (section 5).",
      "**Legend** \u2014 the formatting of each level, for reference.",
      "**Projector** \u2014 the **Presenter view** button (section 7).",
      "**Tracked Changes** \u2014 two buttons that convert between Word\u2019s Tracked Changes and the add-in\u2019s struck-through and underlined text (section 8).",
    ]),
    k.table([1900, 4660, 2800], buttons("[^5]")),

    H1("3. How amendments are displayed"),
    k.legend([900, 2300, 3080, 3080]),
    k.p("All text is black at every level, for easy reading. Each level is shown by its shading (none, light blue or light purple), its number of strike lines, italics at Level 3, and the style and color of its underline (black single, blue double or purple dotted). Strike lines are always black.", { before: 80 }),
    H2("Striking ordinary text"),
    k.p("Striking ordinary text (words that no pending amendment proposes to insert) changes those words\u2019 formatting where they stand. For example, if the main motion would strike \u201cquarterly\u201d and insert \u201cmonthly\u201d, and a primary amendment would strike the ordinary word \u201cPresident\u201d and insert \u201cSecretary\u201d, the screen shows:"),
    k.p("The Board shall meet [[del1:quarterly]] [[ins1:monthly]] at a location designated by the [[del2:President]] [[ins2:Secretary]].", { run: { size: 26 } }),
    H2("Striking text that a lower level proposes to insert"),
    k.p("Struck words are never shown twice. When an amendment proposes to strike words that the level below it proposes to insert, those words are struck where they stand, with the striking level\u2019s strike lines and shading, and they keep the lower level\u2019s underline so a reader can still tell where they came from. For example, if the main motion would insert \u201c$75\u201d, a primary amendment would strike it and insert \u201c$65\u201d, and a secondary amendment would strike \u201c$65\u201d and insert \u201c$60\u201d, the screen shows:"),
    k.p("[[del1:$50]] [[del2u1:$75]] [[del3u2:$65]] [[ins3:$60]]", { run: { size: 26 } }),
    k.p("This works the same however long the struck passage is: a whole sentence the main motion proposes to add, struck by a primary amendment, appears once, double-struck and shaded light blue, still bearing its black underline."),

    H1("4. Marking amendments"),
    H2("When a new amendment is offered"),
    k.p("Click its level in the pane: **2** for a primary amendment, **3** for a secondary amendment, or **1** for a change to the main motion itself. Click it **even if that level is already lit**. Each click tells the add-in that what you mark next is a new amendment."),
    k.p("A secondary amendment needs a pending primary amendment to amend; if an attempt is made to insert one where there is no pending primary amendment, the status box explains the error."),
    H2("Striking words"),
    k.p("Select the words in the document, then click **Strike**. It doesn\u2019t matter whether the selection includes spaces before or after the words."),
    k.p("To strike a whole paragraph, select all of its words\u2014triple-click in the paragraph, or drag across it (see footnote 5)\u2014then click **Strike**. If the paragraph has a section number or letter, it is struck the same way as the words, though without their shading. It doesn\u2019t matter whether the selection includes the paragraph break at the end of the line above or below; the add-in leaves it out. A selection that starts in part of one paragraph and ends in part of another can\u2019t be marked: mark each paragraph\u2019s part separately, or select whole paragraphs."),
    H2("Inserting words"),
    k.p("Type the new words in the document where they belong, select them, then click **Insert**. They take on the level\u2019s formatting, and they are removed if the amendment fails. If nothing is selected when you click Insert, the status box reminds you to type and select the words first."),
    k.p("Striking words that were inserted as part of the same amendment level takes them back out (for example, to correct a typing mistake)."),
    H2("Amendments that only strike, or only insert"),
    k.p("Either is a complete amendment. Mark what it does; it is voted on in the same way."),
    H2("Text with Tracked Changes"),
    k.p("**Strike** and **Insert** do nothing if the selected text has Word Tracked Changes, or if Track Changes is on; the status box says why. Click **Tracked Changes → manual markup** (section 8), then mark the amendment. That button works even while an amendment is pending, so if an amendment was offered before the document was converted, there is no need to undo it first."),
    H2("One motion, several places"),
    k.callout(W, "Important", [
      SEVERAL_PLACES[0],
      "Clicking a level button starts a **new** amendment. If you click it between places, each place becomes a separate amendment, voted on its own, even if \u201cAll places\u201d is clicked (see below).",
    ]),
    k.p("A strike and the replacement words typed right before or after it (with only spaces or punctuation, such as a comma, between) count as one place.", { before: 100 }),

    H1("5. Recording votes"),
    H2("Which amendment a click acts on"),
    k.p("Click inside the text of the amendment that was voted on, then click **Adopted** or **Failed**. The button acts on the most recently offered pending amendment whose text contains the cursor. A cursor right at the edge of an amendment\u2019s text counts as inside it."),
    k.p("An amendment\u2019s struck words (including words a lower level proposes to insert, which it strikes where they stand) and its new words all count as its text."),
    k.p("After any vote, the next mark you make starts a new amendment at the level currently lit in the pane, without the need to click the Active Level again. (That level stays as it was: if 3 is lit when a primary amendment is voted on, the next mark is a secondary amendment, so click **2** first if a new primary amendment is offered.)"),
    H2("What adopted and failed do"),
    k.table([2000, 3680, 3680], RESULTS),
    H2("Amendments marked in several places"),
    k.p("Adopted and Failed by default act **only on the place where the cursor is**. Click inside the amendment in each place in turn and click the button; the status box says how many places are still pending, for example \u201cLevel 3 (secondary amendment) adopted here. 2 other places still pending.\u201d"),
    k.p("To apply one result everywhere at once, click **All places** in the pane (it turns yellow), click inside the amendment\u2019s text in any of its places, then click Adopted or Failed once. All places switches itself off after that one vote; clicking it again before recording a voting result cancels it."),
    H2("The order guard"),
    k.p("A secondary amendment must be voted on before its primary amendment, and a primary amendment before the main motion. If you click Adopted or Failed in an amendment while an amendment to it is still pending **in the same place**, the pane says, for example:"),
    k.p("*Cursor in wrong amendment level: the Level 3 (secondary amendment) to this amendment is still pending. Click inside its text and resolve it first.*"),
    k.p("and nothing changes. Amendments pending in other places, elsewhere in the paragraph or in other paragraphs, do not block recording the vote on an amendment. With All places switched on, a higher-level pending amendment in any of the places blocks recording the result of a vote."),
    H2("The main motion (Level 1)"),
    k.p("Votes on changes proposed in a main motion that is one to *Rescind or Amend Something Previously Adopted* are recorded like this:"),
    ...k.steps([
      "Click **1** for the Active Level in the pane.",
      "Click inside the change\u2019s struck or inserted words.",
      "Click **Adopted** or **Failed**.",
    ]),
    k.p("Only that one change is recorded: the struck and inserted words around the cursor, with only spaces or punctuation (such as a comma) between them. Other changes in the same or other paragraphs are left alone. To record several changes together, select them all first, then click the button."),
    k.p("Level 1 works this way, rather than with All places, because the main motion\u2019s changes were in the document beforehand and were not marked with the add-in, so the add-in cannot tell which changes belong together. Selecting them is how you tell it."),
    k.p("This works only while Level 1 is the active level, so a stray click during a Level 2 or 3 vote can\u2019t change the main motion."),

    H1("6. Mistakes and Ctrl+Z (the \u201cUndo\u201d function)"),
    k.callout(W, "Before pressing Ctrl+Z", [
      "Click just to the **right** of the amendment\u2019s text, not inside it. With the cursor inside an amendment\u2019s text, Ctrl+Z may not undo the add-in\u2019s changes.",
    ]),
    k.gap(40),
    ...k.bullets([
      "**Wrong words marked:** repeatedly press `Ctrl+Z` until the mark is gone, then mark again. The add-in ignores marks that Ctrl+Z has taken out.",
      "**Wrong vote clicked:**[^6] repeatedly press `Ctrl+Z` until the text looks exactly as it did before the vote. The amendment is then pending again, and you can record the result of a vote on it as usual. The add-in notices the undo at your next click.",
    ]),
    k.p("Word may record one action of the add-in as several undo steps, so more than one Ctrl+Z is often needed. The text can look odd between steps; keep pressing until it is as before."),

    H1("7. Presenter view: the projected screen"),
    k.p("Presenter view opens a separate window for the screen the members see. It shows only the document, in large type, with each level\u2019s formatting and a small legend along the bottom; no ribbon, pane or page edges. It follows your work: the paragraph containing the cursor stays in the middle of the window, marked by a thin gray bar, and the window updates immediately after every click and within a second of typing."),
    H2("Setting up the screens"),
    k.p("The computer needs a second screen (either a monitor or a digital projector) set to **extend** the desktop, not duplicate it: press **Windows key + P** and choose **Extend**. Where a projector is used to show the second monitor\u2019s picture, that pair can be duplicated: search Windows for \u201cDuplicate or extend to a connected display\u201d and, when it comes up, click **Identify** to see the numbers assigned to the second monitor and the projector, then use the drop-down to choose the setting that duplicates those two displays."),
    H2("Opening it and moving it to the projected screen"),
    ...k.steps([
      "Click **Presenter view** in the pane\u2019s **Projector** section. The window opens, filling the screen Word is on.",
      "Click the window\u2019s **title bar** (the thin strip along its top) once.",
      "Press **Windows key + Shift + \u2192** (or **\u2190** if the projected screen is on the left). The whole window moves to the other screen, keeping its size. (See below if the moved window does not fill the other screen.)",
      "Click back in the Word document (on the original computer monitor) to carry on working.",
    ]),
    H2("Making it fit the projected screen"),
    k.p("Word\u2019s add-in windows have no maximize button, so if the two screens differ in size the window may not fit the projected screen exactly. Adjust it once:"),
    ...k.bullets([
      "**To move the window:** drag its title bar until its top-left corner is at the screen\u2019s top-left corner.",
      "**To resize it:** move the mouse pointer to the window\u2019s right or bottom edge until the pointer becomes a two-headed arrow, then drag the edge to the edge of the screen. Dragging the **bottom-right corner** adjusts both at once.",
      "**If part of the window is off the screen:** move it first so its top-left corner shows, then drag the bottom-right corner in.",
    ]),
    H2("While presenting"),
    ...k.bullets([
      "Make all amendment changes in Word using the original screen; do not attempt to make any directly on the projected screen.",
      "To show a different part of the document, click there in Word; the window on the other screen follows.",
      "**Footnotes and endnotes:** click in a footnote or endnote in Word, and the window shows the paragraph containing that note’s number (highlighted), with the note in a shaded panel below it, both with their amendments. If both don’t fit, the window keeps the part of the note you are working in visible; **A−** makes room for both. Click back in the main text to return to it. Switching back and forth only changes what the window shows; pending amendments in the note and in the main text are unaffected.",
      "**Text size and legend:** move the mouse over the presenter window in the projected screen and controls appear in its top-right corner: **A\u2212** decreases and **A+** increases the text size (remembered for next time), and **Legend** shows or hides the legend. With the presenter window active, the **\u2212**, **+** and **L** keys do the same.",
      "**To close it:** click **Presenter view** in the add-in pane again, or the window\u2019s **X**.",
    ]),
    k.p("The window shows the document\u2019s text, not Word\u2019s page layout: headings, bold, italics, tables and footnote and endnote numbers come through, but page breaks, headers and footers, pictures and Word\u2019s automatic numbering do not. The text of a footnote or endnote appears only while the cursor is in it (see above). Typed numbers appear as usual."),

    H1("8. Converting Tracked Changes"),
    k.p("A motion is sometimes prepared in Word with **Track Changes** on, so that its proposed changes appear as Word’s Tracked Changes rather than as struck-through and underlined text. The add-in can’t show amendments to Tracked Changes, so convert them beforehand. After the votes, the result can be converted back into Tracked Changes, for example for the minutes or for whoever prepares the final text. The two buttons for this are in the pane’s **Tracked Changes** section, below Presenter view."),
    k.callout(W, "Before converting", [
      "{{warn:Make a copy of the document first (File → Save As), and convert the copy, not the original!}}",
      "The buttons act **at once, with no confirmation**, and a conversion changes many places at once. Ctrl+Z can undo it, but may take many presses.",
      "Converting **Tracked Changes → manual markup** is especially hard to undo. Converting the result back with the other button does not restore the original: every change is then recorded as made by you, today, rather than by its original author on its original date; any tracked change of formatting is lost; and any other underlined or struck-through text in the document becomes a tracked change too. The copy of the original is the only sure way back.",
      "**With text selected**, a button converts only the selected text, whether in the main text or in a footnote or endnote. **With nothing selected** (just a cursor), it converts the whole document, including every footnote and endnote.",
    ]),
    H2("Beforehand: Tracked Changes → manual markup"),
    ...k.steps([
      "Make a copy of the document and open the copy (see the box above).",
      "Select the text to convert or, to convert all Tracked Changes in the document, click anywhere in the document without selecting anything.",
      "Click **Tracked Changes → manual markup**. The status box says how many changes were converted.",
    ]),
    k.p("Each tracked deletion becomes [[del1:struck-through]] text and each tracked insertion becomes [[ins1:underlined]] text in the main motion’s (Level 1) formatting, which the add-in reads as the main motion’s proposed changes (section 5). Any other kind of tracked change, such as a change of formatting, is accepted. The Track Changes feature, if it is on, is turned off.", { before: 80 }),
    H2("After the votes: manual markup → Tracked Changes"),
    k.p("Occasionally, as sometimes with a motion to Rescind or Amend Something Previously Adopted, it may be desired to show the adopted amendments as Tracked Changes. Before clicking **Manual markup → Tracked Changes** to do so, be sure to record the result of the vote on every pending primary and secondary amendment first. If any amendment in the selected text (or, with nothing selected, anywhere in the document) is still pending, including one in a footnote or endnote, the status box says *“Votes must be recorded on pending amendments before converting to Tracked Changes”* and nothing changes."),
    ...k.steps([
      "Select the text to convert or, to convert all amendments shown in the document, click anywhere in the document without selecting anything.",
      "Click **Manual markup → Tracked Changes**. The status box says how many passages were converted.",
    ]),
    k.p("Struck-through text becomes tracked deletions and underlined text becomes tracked insertions, shown in Word’s usual Track Changes colors. Track Changes is left **on**, so later edits are tracked too, and the add-in will not mark amendments in that document until it is turned off (**Review → Track Changes**). Do not use the **Tracked Changes → manual markup** button for that: it would turn the new Tracked Changes back into manual markup.", { before: 80 }),
    H2("Points to note"),
    ...k.bullets([
      "Only single underlining and single strikethrough are converted. Any such text is converted, even if it is underlined or struck through for some other reason (section 14).",
      "Amendments of the kinds this add-in handles (inserting words, striking words, or striking and inserting words) convert cleanly. A change that strikes or inserts whole paragraphs may not.",
      "If struck-through text includes a footnote or endnote number, Word shows that whole note as deleted, as it always does when a note’s number is deleted with Track Changes on. Any changes inside the note are converted first and are not lost: rejecting the changes brings the note back as it was, and accepting them removes it. To avoid striking an entire footnote or endnote, attach the number for that note to text that is not stricken.",
      "The buttons need Word version 2308 or later; older versions say so in the status box (section 12).",
    ]),

    H1("9. Saving, closing and reopening"),
    k.p("Everything the add-in knows about pending amendments is kept inside the document itself. **Save** the document (`Ctrl+S`, or click “File” on the Word ribbon and then “Save” on the list that appears) to keep it, for example before a recess or when a matter is postponed or laid on the table. When the saved document is reopened, the add-in pane opens by itself and votes on every pending amendment can be recorded exactly as before. Presenter view is not saved; open it again when needed."),
    k.p("Closing without saving discards every mark and vote since the last save, as with any Word edit. For practice runs, this is a convenient way to keep the original document unchanged."),

    H1("10. Messages in the status box"),
    k.table([3400, 2980, 2980], [
      ["Message", "What it means", "What to do"],
      ["Cursor in wrong amendment level: the Level \u2026 to this amendment / passage / change is still pending\u2026", "An amendment to the one at the cursor is still pending in this place.", "Click inside the text of the level it names and record the result of the vote on that first."],
      ["Cursor is not in the immediately pending amendment\u2026", "The cursor isn\u2019t in any pending amendment\u2019s text.", "Click inside the text of the amendment it names, then click the button again."],
      ["No amendment is pending.", "Every amendment has been voted on.", "Nothing; or click 1 to change the main motion."],
      ["No Level 1 change at the cursor\u2026", "Level 1 is active, but the cursor is in words that do not show any proposed amendment.", "Click inside words that are struck or underlined (inserted)."],
      ["Select the words to strike first, then click Strike.", "Strike was clicked with nothing selected.", "Select the words, then click Strike."],
      ["Type the words first, select them, then click Insert.", "Insert was clicked with nothing selected.", "Type the new words in the document, select them, then click Insert."],
      ["There is no pending Level 2 (primary) amendment for this secondary amendment to amend.", "Level 3 is active but no primary amendment is pending.", "Mark the primary amendment first (click 2)."],
      ["Word couldn\u2019t mark that selection\u2026", "Word refused to mark the selection, usually because it partly overlaps another amendment\u2019s text or covers parts of more than one paragraph.", "Keep the selection within one paragraph (or select whole paragraphs), and select only inside, or only outside, the other amendment’s text."],
      ["That selection covers parts of more than one paragraph\u2026", "The selection runs from the middle of one paragraph into another.", "Mark each paragraph\u2019s part separately, or select whole paragraphs (section 4)."],
      ["Only a paragraph break is selected\u2026", "The selection holds no words, only the break between two paragraphs.", "Select the words, then click the button again."],
      ["Another pending amendment is marked inside this one\u2026", "Another amendment\u2019s text sits inside this one.", "Vote on the inner one first."],
      ["Presenter view is open. Click its title bar\u2026", "The presenter window has opened.", "Move it to the projected screen (section 7)."],
      ["Presenter view couldn\u2019t open\u2026 / This version of Word can\u2019t run Presenter view\u2026", "Word refused to open the window, or is too old for it.", "Close and reopen the pane; check Word\u2019s version (section 12)."],
      ["Converted in the whole document / the selected text: \u2026", "The conversion is done; the message gives the numbers converted.", "Nothing; check the result."],
      ["There are no Tracked Changes in \u2026 / There is no struck-through or underlined text in \u2026", "There was nothing to convert where the button looked. If Track Changes was on, Tracked Changes \u2192 manual markup has turned it off, and says so.", "To convert the whole document, click in it without selecting anything, then click the button again."],
      ["Track Changes is on, so Word would record the amendment as Tracked Changes\u2026", "Strike or Insert was clicked while Word\u2019s Track Changes is on.", "Click **Tracked Changes \u2192 manual markup**, which converts any Tracked Changes and turns Track Changes off; then mark the amendment (section 4)."],
      ["This text has Tracked Changes\u2026", "Strike or Insert was clicked on text that has Word Tracked Changes.", "Click **Tracked Changes \u2192 manual markup**, then mark the amendment (section 4)."],
      ["Votes must be recorded on pending amendments before converting to Tracked Changes.", "Manual markup \u2192 Tracked Changes was clicked while an amendment in the selected text (or, with nothing selected, in the document) still awaits its vote.", "Record the result of the vote on each pending amendment, then click the button again (section 8)."],
      ["This version of Word can\u2019t convert \u2026", "Word is too old for the Tracked Changes buttons.", "Use Word version 2308 or later (section 12)."],
      ["Word couldn\u2019t finish the conversion: \u2026", "Word refused part of the conversion.", "Press Ctrl+Z until the text is as before, or close the document without saving; note the message and tell whoever maintains the add-in."],
      ["Something went wrong: \u2026", "An unexpected error.", "Press Ctrl+Z if the text changed; note the message and tell whoever maintains the add-in."],
    ]),

    H1("11. Troubleshooting"),
    k.table([2900, 6460], [
      ["Problem", "What to do"],
      ["Amendment Layers isn\u2019t on the Home tab", "Look under **Home \u2192 Add-ins**. If it isn\u2019t there, it isn\u2019t installed on this computer (section 12)."],
      ["The pane says \u201cLoading\u2026\u201d forever, or shows an error about a certificate or blocked content", `Check the internet connection: **${SITE}** should open in a web browser. Then close the add-in pane with its **X** and reopen it from the Home tab.`],
      ["\u201cADD-IN ERROR This add-in is no longer available\u2026 only available during debugging\u2026\u201d", "Word\u2019s add-in cache is out of date. Close Word completely. In File Explorer\u2019s address bar, type **%LOCALAPPDATA%\\Microsoft\\Office\\16.0** and press Enter. In that folder, rename the folder **Wef** (for example to **Wef.old**), and reopen Word. Word creates a new **Wef** folder."],
      ["Gray \u201cClick or tap here to enter text.\u201d appears", "Click any of the add-in\u2019s buttons: the add-in removes these words after every action."],
      ["Ctrl+Z seems to do nothing", "Click just to the right of the amendment\u2019s text (not inside it), then press Ctrl+Z again."],
      ["A vote did the wrong thing", "Click just to the right of the amendment\u2019s text and press `Ctrl+Z` until the text is as before; then vote again (section 6)."],
      ["Presenter view opened on the wrong screen, or doesn\u2019t fit", "Move and resize it as in section 7. Windows key + Shift + \u2192 or \u2190 moves it between screens."],
      ["Presenter view shows \u201cWaiting for Word\u2026\u201d or stops following", "Click in the Word document. If it still doesn\u2019t update, close it with the Presenter view button and open it again."],
      ["Keyboard shortcuts do nothing", "Click in the document text first; the shortcuts don\u2019t work while the add-in pane has focus. Otherwise use the button, and see section 13."],
      ["The document has Tracked Changes", "Convert them, before amendments are marked, with **Tracked Changes \u2192 manual markup** (section 8)."],
      ["A Tracked Changes conversion was done by mistake", "Press `Ctrl+Z` until the text is as before, or close the document without saving. Converting back with the other button does not fully restore Tracked Changes (section 8); go back to the copy of the original instead."],
      ["An amendment is offered to text that is still in Tracked Changes", "Click **Tracked Changes \u2192 manual markup**, then mark the amendment. If an amendment was already marked, there is no need to undo it first (section 4)."],
      ["After converting to Tracked Changes, Strike and Insert say Track Changes is on", "The conversion leaves Track Changes on. To mark further amendments, turn it off with **Review \u2192 Track Changes** (not with the other conversion button)."],
    ]),

    H1("12. Installing on a computer"),
    H2("What is needed"),
    ...k.bullets([
      "Word for Windows, version 2408 or later, either in **Microsoft 365** (subscription) or in **Office 2021 or Office 2024 bought outright** (\u201cretail\u201d editions such as Office Home & Business) kept up to date. Retail editions receive the same updates as Microsoft 365; version 2608, for example, is one of them.",
      "**Office LTSC 2024** (the volume-licensed edition an organization\u2019s IT department installs) also works, except for the keyboard shortcuts: use the pane\u2019s buttons.",
      "To check the edition and version: **File \u2192 Account**, then **About Word**.",
      "The add-in **can\u2019t be installed** on Office 2019, Office 2016 or Office LTSC 2021: their Word lacks features the add-in needs. Word for Mac and Word on the web support the add-in\u2019s features, but these instructions cover Windows only.",
      "An internet connection at the meeting.",
      "A web browser, to download the add-in\u2019s file, **manifest.xml**. No Netlify account is needed.",
    ]),
    H2("Step A \u2014 download the add-in\u2019s file"),
    ...k.steps([
      `In a web browser, go to **${SITE}/install.html**.`,
      "Click **Download manifest.xml**. The browser saves **manifest.xml**, normally in your **Downloads** folder. If the browser asks whether to keep the file, choose **Keep**.",
    ]),
    H2("Step B \u2014 a shared folder for the file (once)"),
    ...k.steps([
      "Create a folder, for example **C:\\AmendmentLayers**, and move **manifest.xml** from your Downloads folder into it. (An existing network share works too; then every computer can use the same folder.) Do not use a folder inside OneDrive, Google Drive, Dropbox or another cloud-synced folder; Word needs the file to be on the computer itself (or on the network share). Windows often keeps the Desktop and Documents folders in OneDrive, so a folder directly on **C:\\** is safest.",
      "Right-click the folder \u2192 **Properties \u2192 Sharing \u2192 Share\u2026**, then **Share**. Note the network path shown, such as **\\\\COMPUTERNAME\\AmendmentLayers**.",
    ]),
    H2("Step C \u2014 in Word, on each computer"),
    ...k.steps([
      "**File \u2192 Options \u2192 Trust Center \u2192 Trust Center Settings\u2026 \u2192 Trusted Add-in Catalogs.**",
      "In **Catalog Url**, type the network path from step B, click **Add catalog**, tick **Show in Menu**, then **OK** and **OK**.",
      "Close Word completely and reopen it.",
      "**Home \u2192 Add-ins \u2192 More Add-ins** (or **Insert \u2192 Get Add-ins**), choose the **SHARED FOLDER** tab, select **Amendment Layers**, and click **Add**.",
      "The **Amendment Layers** button now appears on the Home tab.",
    ]),
    k.p("Step C changes a Word security setting: it tells Word to trust the add-ins in that folder. Only add a folder you control.", { before: 80 }),
    H2("Updates"),
    k.p(`Changes to how the add-in works reach every computer automatically the next time the pane opens. Only a change to **manifest.xml** itself (rare) needs the new file downloaded again from **${SITE}/install.html** and copied into the shared folder in place of the old one; then close and reopen Word.`),

    H1("13. Keyboard shortcuts"),
    k.p("Each of the four main buttons also has a keyboard shortcut (a set of keys to be pressed simultaneously) for anyone who prefers the keyboard. Click in the document text first: the shortcuts don\u2019t work while the add-in pane has focus. They are not available in Office LTSC 2024 (section 12)."),
    k.table([2700, W - 2700], [
      ["Keys", "Same as"],
      ["`Ctrl+Alt+Shift+D`", "**Strike**"],
      ["`Ctrl+Alt+Shift+I`", "**Insert**. With nothing selected, these keys also let you insert by typing: press them, then type the words straight into the document."],
      ["`Ctrl+Alt+Shift+Y`", "**Adopted** (\u201cyes\u201d)"],
      ["`Ctrl+Alt+Shift+N`", "**Failed** (\u201cno\u201d)"],
    ]),
    k.p("The first time the keys are pressed, Word may ask which action to use; choose the Amendment Layers action. To change that choice later, type **Reset Office Add-ins shortcut preferences** in Word\u2019s search box. If a set of keys does nothing on a particular computer, another program there is probably using that combination; ask whoever maintains the add-in for a different one.", { before: 80 }),

    H1("14. Limits of this version"),
    ...k.bullets([
      "**Ctrl+Z** works with the cursor just outside an amendment\u2019s text, not inside it (section 6).",
      "**A selection that partly overlaps** another amendment\u2019s text can\u2019t be marked. Select wholly inside it or wholly outside it.",
      "**An amendment spanning more than one paragraph** must cover whole paragraphs. A selection that starts in part of one paragraph and ends in part of another can’t be marked; the status box says so (section 4).",
      "**A struck paragraph’s section number or letter** is shown with strikeout lines like its words but not shaded. If the document’s **last** paragraph is struck in full and the strike is finally adopted, Word keeps an empty line in its place, without a number.",
      "**Main motion:** the add-in recognizes the main motion\u2019s changes only by their underlining and strikethrough. If the document also uses underlining or strikethrough for something else (for example, an underlined heading or defined term), the add-in treats that text as a Level 1 change too: recording a vote on a neighboring change could remove that underline or, on Failed, delete those words. Avoid underlining or striking through text in the document except to show the main motion\u2019s proposed changes.",
      "**All places** applies to amendments marked with the add-in. For the main motion, select the changes instead (section 5).",
      "**Presenter view** shows text rather than Word\u2019s page layout, and keeps a thin title bar (section 7).",
      "**Internet needed:** the pane and Presenter view load from the web.",
      "**Tracked Changes conversion** treats all single underlining and strikethrough as proposed changes, and may not convert a change spanning paragraphs cleanly (section 8).",
    ]),

    H1("15. For whoever maintains the add-in"),
    k.p(`Technical details (hosting, publishing updates, the code and its tests) are in **README.md** in the add-in’s project folder (currently **${PROJECT}** on the developer’s computer). The web files are published at **${SITE}**.`),
  ];

  return new Document({
    creator: "Amendment Layers",
    title: "Amendment Layers \u2014 User Manual",
    styles: styles(22, 30, 25),
    numbering: numbering(),
    footnotes: footnotes(MANUAL_FOOTNOTES),
    sections: [{
      properties: { page: { size: LETTER, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      footers: { default: footer("Amendment Layers \u2014 User Manual") },
      children,
    }],
  });
}

// ---------------------------------------------------------------------- build

async function write(doc, name) {
  const file = path.join(__dirname, name);
  fs.writeFileSync(file, await Packer.toBuffer(doc));
  console.log(`Wrote ${file}`);
  return file;
}

// Word fills in the table of contents; until then it is empty.
function updateToc(file) {
  try {
    execFileSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass",
      "-File", path.join(__dirname, "update-toc.ps1"), file], { stdio: "inherit" });
    console.log("Filled in the Contents (page numbers and links).");
  } catch (e) {
    console.warn("Could not have Word fill in the Contents. Is the file open in Word? "
      + "Close it and run this again, or in Word right-click the Contents and choose Update Field.");
  }
}

(async () => {
  await write(quickReference(), "Amendment Layers - Quick Reference.docx");
  updateToc(await write(manual(), "Amendment Layers - User Manual.docx"));
})();
