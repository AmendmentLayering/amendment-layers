/*
 * Builds "Amendment Layers - Word on the web test plan.docx": what to test
 * and review before the User Manual covers Word on the web.
 * Run: node docs/build-web-test.js
 */
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, LevelFormat, Packer, PageNumber, Paragraph,
  ShadingType, Table, TableCell, TableRow, TextRun, WidthType, HeadingLevel,
} = require("docx");

const FONT = "Calibri";
const WARN_RED = "C00000";
const SITE = "https://amendment-layers.netlify.app";
const OUT = path.join(__dirname, "Amendment Layers - Word on the web test plan.docx");
const WIDTH = 9360; // 6.5" of text width

// **bold**, *italic*, {{warn:...}} (bold italic red)
function runs(text, base = {}) {
  const out = [];
  const re = /\*\*(.+?)\*\*|\*(.+?)\*|\{\{warn:(.+?)\}\}|([^*{]+|[*{])/g;
  let m;
  while ((m = re.exec(text))) {
    if (m[1]) out.push(new TextRun({ text: m[1], bold: true, font: FONT, ...base }));
    else if (m[2]) out.push(new TextRun({ text: m[2], italics: true, font: FONT, ...base }));
    else if (m[3]) out.push(new TextRun({ text: m[3], bold: true, italics: true, color: WARN_RED, font: FONT, ...base }));
    else out.push(new TextRun({ text: m[4], font: FONT, ...base }));
  }
  return out;
}

const p = (text, opts = {}) => new Paragraph({ children: runs(text), spacing: { after: 120, ...(opts.spacing || {}) }, ...opts.para });
const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text, font: FONT })], spacing: { before: 280, after: 120 }, keepNext: true });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text, font: FONT })], spacing: { before: 200, after: 80 }, keepNext: true });
let listN = 0;
const steps = (items) => {
  const ref = `steps${listN++}`;
  numbering.push({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 360 } } } }] });
  return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 60 } }));
};
const bullets = (items) => items.map((t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: runs(t), spacing: { after: 60 } }));
const numbering = [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 360 } } } }] }];

const border = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, width, head) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: head ? { fill: "E7E6E6", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: String(text).split("\n").map((line) => new Paragraph({ children: runs(line, head ? { bold: true } : {}), spacing: { after: 40 } })),
  });
}
// Test table: number, what to do, what should happen, result (left blank).
let testNo = 0;
function tests(rows) {
  const w = [500, 3700, 3260, 1900];
  return new Table({
    width: { size: WIDTH, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      new TableRow({ tableHeader: true, children: ["#", "Do this", "What should happen", "Result / notes"].map((t, i) => cell(t, w[i], true)) }),
      ...rows.map(([doThis, expect]) => new TableRow({
        cantSplit: true,
        children: [cell(String(++testNo), w[0]), cell(doThis, w[1]), cell(expect, w[2]), cell("", w[3])],
      })),
    ],
  });
}
const gap = () => new Paragraph({ children: [], spacing: { after: 60 } });

const body = [
  new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: "Amendment Layers — Word on the web test plan", font: FONT })] }),
  p("The add-in runs the same code in Word on the web as in Word for Windows, so this is mostly a check that each feature behaves the same, plus a review of draft instructions for the User Manual. Part A sets up; Part B is the tests; Part C is the draft instructions to review. Expect it to take one to two hours."),
  p("For each test, note in the last column **OK**, or what happened instead (the exact wording of any message in the status box helps most). Tests marked **(web)** are where Word on the web is most likely to differ from Windows."),

  h1("Part A — Setting up"),
  h2("What is needed"),
  ...bullets([
    "A Microsoft account (the one used for a Microsoft 365 subscription works; Word on the web also works with a free account).",
    "Microsoft Edge or Google Chrome on a Windows computer, ideally with a second screen for the Presenter view tests.",
    "Copies of two test documents: one with struck-through and underlined text including a footnote and an endnote (such as the one used for the Tracked Changes tests), and any short document with a main motion to amend.",
  ]),
  h2("Putting the test documents in OneDrive"),
  p("Word on the web opens documents stored in OneDrive. Nothing needs installing, and Google Drive is not affected."),
  ...steps([
    "In the browser, go to **onedrive.live.com** and sign in with the Microsoft account.",
    "Drag the test documents from their folder onto the OneDrive page. This uploads copies; the originals stay where they are.",
    "Click a document to open it in Word on the web. If it opens in a read-only view, choose **Editing** (or **Edit Document → Edit in Browser**).",
    "To keep a result, use **File → Save As → Download a Copy** in Word on the web.",
  ]),
  h2("Adding the add-in"),
  ...steps([
    `In a new browser tab, go to **${SITE}/install.html** and click **Download manifest.xml**. (Ignore the page’s instructions after that; they are for Windows.)`,
    "Back in the document in Word on the web: **Home → Add-ins**, then **More Settings** (or **Advanced…**).",
    "In the **Office Add-ins** box, choose **Upload My Add-in**, click **Browse**, choose **manifest.xml** from the Downloads folder, then **Upload**.",
  ]),
  p("Note where the **Amendment Layers** button appears (on the Home tab, at its right end, or only under **Add-ins**); test 2 asks.", { spacing: { before: 80 } }),

  h1("Part B — Tests"),
  h2("Opening the pane"),
  tests([
    ["Upload the add-in as in Part A.", "No error. **Amendment Layers** appears on the ribbon or in the Add-ins menu."],
    ["Click **Amendment Layers**.", "The pane opens at the side, with the level buttons **1 2 3**, **Strike**, **Insert**, **Adopted**, **Failed**, **All places**, **Presenter view**, the two Tracked Changes buttons and the legend. Note where the button was found."],
  ]),
  gap(),
  h2("Marking amendments"),
  tests([
    ["Click **2**. Select a few words of the main motion and click **Strike**.", "The words are double struck through, shaded light blue, all text black, as on Windows."],
    ["Put the cursor just after the struck words, click **Insert** and type replacement words.", "The new words have a blue double underline and light blue shading."],
    ["Click **3**. Select one of the words just inserted and click **Strike**; then click **Insert** and type a replacement. **(web)**", "Level 3 formatting: single strike or purple dotted underline, italics, light purple shading."],
    ["Look closely around each amendment. **(web)**", "No boxes, brackets, tags or colored outlines around the amendments (the add-in’s markers are meant to be hidden)."],
  ]),
  gap(),
  h2("Recording votes"),
  tests([
    ["Click inside the primary amendment’s words (not the secondary’s) and click **Adopted**.", "Refused: the status box says “Cursor in wrong amendment level…” and nothing changes."],
    ["Click inside the secondary amendment’s words and click **Adopted** (or **Failed**).", "The secondary is resolved as on Windows; the primary is still shown."],
    ["Click inside the primary and click **Failed** (or **Adopted**).", "The primary is resolved; the text returns to plain (Failed) or takes the amendment (Adopted)."],
    ["Press **Ctrl+Z** a few times.", "Each vote and marking is undone in turn, as on Windows."],
    ["Mark one primary amendment in two places (click **2** once, strike in both places). Click **All places**, then in one place click **Adopted**.", "Both places are resolved by the one click; **All places** switches itself off."],
    ["In a document whose main motion already has underlined and struck text, click **1**, click in one change and click **Adopted**.", "Only that change is resolved."],
  ]),
  gap(),
  h2("Keyboard shortcuts (web)"),
  p("In Word on the web the shortcuts only work while the cursor is in the document: click in the document first, not in the pane."),
  tests([
    ["Click **2**, then click in the document and select words; press **Ctrl+Alt+Shift+D**.", "The words are struck at Level 2. If the browser does something else instead (opens a menu, a new tab, etc.), note what."],
    ["Put the cursor where words should go; press **Ctrl+Alt+Shift+I** and type.", "The new words are inserted at Level 2."],
    ["Click in the amendment; press **Ctrl+Alt+Shift+Y** (adopted). Mark another and press **Ctrl+Alt+Shift+N** (failed).", "Each is resolved, as with the buttons."],
    ["Click in the pane, then press **Ctrl+Alt+Shift+D** with words selected.", "Probably nothing happens (expected on the web). Note what does."],
    ["If Word asks which action a shortcut should run, note the wording and choose **Amendment Layers**.", "The choice is remembered."],
  ]),
  gap(),
  h2("Presenter view (web)"),
  tests([
    ["Click **Presenter view**.", "A separate browser window opens showing the passage around the cursor in large type, with the legend. If the browser blocks it as a pop-up, allow pop-ups for this site, click **Presenter view** twice (close, open) and note what the browser showed."],
    ["Move the window to the second screen: click its title bar and press **Windows key + Shift + →** (or drag it).", "It moves to the other screen."],
    ["Try to make it fill that screen: its **maximize** button if it has one, double-clicking its title bar, or **F11**.", "Note which works, and whether the address bar or browser toolbars stay visible."],
    ["In Word, click in different paragraphs, mark an amendment, and type a few words.", "The window follows the cursor (thin gray bar), and updates after each click and within about a second of typing."],
    ["Hover over the window: try **A−**, **A+** and **Legend**; with the window active, press **−**, **+** and **L**.", "The text size and legend change. Close and reopen the window: the size is remembered."],
    ["Click in a footnote or endnote in Word (if Word on the web shows them where you can click).", "The window shows the note’s text while the cursor is in it."],
    ["Look at the window’s title or address bar.", "It includes the document’s file name."],
    ["Close it with **Presenter view** in the pane, then reopen it and close it with its **X**.", "Both close it; the pane’s button shows it closed."],
  ]),
  gap(),
  h2("Saving and reopening (web)"),
  tests([
    ["Mark a primary amendment and leave it pending. Wait a few seconds, then close the browser tab.", "Word on the web saves automatically."],
    ["Reopen the document from OneDrive.", "The pane reopens by itself (note if it doesn’t; then click **Amendment Layers**)."],
    ["Click in the pending amendment and click **Adopted** or **Failed**.", "It is resolved as if never closed."],
  ]),
  gap(),
  h2("Tracked Changes"),
  p("{{warn:Use a fresh copy of the test document for each conversion!}}"),
  tests([
    ["With nothing selected, click **Manual markup → Tracked Changes**.", "The status box gives the same counts as in Word for Windows (7 struck, 5 underlined for the earlier test document). Struck and underlined text become Tracked Changes, footnote and endnote included, each note with its number once."],
    ["With Track Changes still on, select words and click **Strike**.", "Refused: “Track Changes is on, so Word would record the amendment as Tracked Changes…”"],
    ["Click **Tracked Changes → manual markup**.", "The Tracked Changes become struck and underlined text again; the status box says Track Changes is off. **Strike** then works."],
    ["In a document with Tracked Changes and Track Changes off (turn it off on the **Review** tab), select text including a Tracked Change and click **Strike**.", "Refused: “This text has Tracked Changes…”"],
  ]),
  gap(),
  h2("General"),
  tests([
    ["Throughout: note anything that took noticeably longer than on Windows (more than about two seconds after a click).", "Clicks take effect within a second or two."],
    ["Note any message in the status box not covered above, with its exact wording.", "—"],
  ]),

  h1("Part C — Draft instructions to review"),
  p("If the tests go well, the User Manual would gain the following for Word on the web. Please mark anything unclear, wrong or missing (for example with Track Changes or comments in a copy of this document)."),
  h2("Installing in Word on the web"),
  ...steps([
    "Word on the web works on documents stored in OneDrive (or SharePoint). If the document is elsewhere, upload a copy: go to **onedrive.live.com**, sign in, and drag the document onto the page.",
    "Open the document in Word on the web, in editing mode.",
    `In another browser tab, go to **${SITE}/install.html** and click **Download manifest.xml**.`,
    "Back in the document: **Home → Add-ins → More Settings → Upload My Add-in**, then **Browse**, choose **manifest.xml**, and **Upload**.",
    "Click **Amendment Layers** on the Home tab to open the pane.",
  ]),
  p("The add-in is remembered by that browser on that computer. In another browser, on another computer, or after the browser’s stored data is cleared, upload **manifest.xml** again. (An organization with Microsoft 365 for business can instead have its administrator add the add-in for everyone.)", { spacing: { before: 80 } }),
  h2("Differences from Word for Windows"),
  ...bullets([
    "**Keyboard shortcuts** are the same, but work only while the cursor is in the document, not while the pane is selected. Click in the document first.",
    "**Presenter view** opens as a separate browser window. If the browser blocks it as a pop-up, allow pop-ups for the site and click **Presenter view** again. Move it to the projected screen by dragging it, or with **Windows key + Shift + →**; [to be completed from tests 19–20: how to make it fill the screen].",
    "**Saving** is automatic. Close the browser tab when finished; the pending amendments are kept, as on Windows.",
    "Word on the web is always up to date, so the version requirements in section 12 don’t apply.",
  ]),

  h1("What to send back"),
  p("This document with the Result column filled in and any marks on Part C, plus screenshots of anything that looked wrong (Windows key + Shift + S takes one), and a copy of any test document that came out wrongly (**File → Save As → Download a Copy**)."),
];

const doc = new Document({
  creator: "Amendment Layers",
  title: "Amendment Layers — Word on the web test plan",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { size: 36, bold: true, font: FONT }, paragraph: { spacing: { after: 160 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: FONT, color: "1F3864" }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 23, bold: true, font: FONT, color: "2F5496" }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: numbering },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", PageNumber.CURRENT], font: FONT, size: 18 })] })] }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("Wrote " + OUT);
});
