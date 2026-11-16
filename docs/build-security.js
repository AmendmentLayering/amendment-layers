// Builds the cybersecurity vetting document from the one text below:
//   ../SECURITY.md   (for the GitHub repository; GitHub links it as the
//                     repository's "Security policy")
//   Cybersecurity Vetting of the Amendment Layers Word Add-in.docx
//                    (a Word copy for review and for sharing with clients)
// Edit the text here, then run:  node docs/build-security.js
//
// In the text, **words** are bold. A paragraph given as an array is a
// bulleted list; one given as { quote: "..." } is an indented quotation.
//
// Keep this document true to the code: check it whenever the add-in changes
// what it loads, sends or stores.

const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber } = require("docx");

const TITLE = "Cybersecurity Vetting of the Amendment Layers Word Add-in";
const REPO = "https://github.com/AmendmentLayering/amendment-layers";
const SITE = "https://amendment-layers.netlify.app";

const SECTIONS = [
  [null, [
    "This document gives an organization’s cybersecurity team the information it needs to vet the Amendment Layers Word add-in.",
  ]],
  ["Where to find the code", [
    "As explained in section 15 of the Amendment Layers User Manual (Technical details and license):",
    { quote: `The complete project, including the code, its tests and other technical files, is publicly available at **${REPO}**, for example for cybersecurity vetting. Technical details (hosting, publishing updates, the code and its tests) are in its **README.md** file. The add-in may be used under the terms in its **LICENSE.md** file.` },
    { quote: `The web files that Word runs are published at **${SITE}**. The project’s **check-live.ps1** script confirms that the files published there are the same as those in the project.` },
    "The **check-live.ps1** script also accepts a **-Site** setting naming a different web address. A team that hosts its own copy of the add-in (see below) can use it to confirm at any time that the files on its own server still match the version it vetted.",
  ]],
  ["What the add-in sends and stores", [
    [
      "The add-in never sends a document’s contents anywhere. Its only network requests download the add-in’s own files from wherever it is hosted, and Microsoft’s add-in library (described below).",
      "Opening Presenter view also includes the document’s file name in that window’s web address, so that the window’s title bar names the document. The file name may therefore appear in the hosting server’s request logs. Only the file name is included, never the document’s contents.",
      "A record of amendments that are proposed to a Word document, with their disposition, as well as brief diagnostic notes, is stored inside the Word document itself. Presenter view also remembers its text size and legend setting on the computer.",
    ],
  ]],
  ["Avoiding the prospect that later changes could introduce spyware or other malware", [
    "As noted in section 12 of the User Manual under Updates, “Changes to how the add-in works reach every computer automatically the next time the pane opens.” An organization may be concerned that, even after the add-in’s code has been vetted, spyware or other malware could be introduced into code downloaded afterwards. There is a remedy: the cybersecurity team can host the add-in on its own internal server and install it from there. Then nothing is downloaded from outside except Microsoft’s standard add-in library, and updates reach the organization only when the cybersecurity team chooses to vet and install them. No special version needs to be prepared: the team uses the same files that are publicly available.",
    "Word add-ins of this type are small web pages that Word displays inside its own window. Word always loads them from a web address given in the add-in’s **manifest.xml** file; it cannot run them from ordinary files on disk. So the add-in cannot be made into a traditional program installed from a disk. But the web address does not have to be Netlify or anywhere on the public internet.",
    "**Points the cybersecurity team will want to know:**",
    [
      "One outside component remains: Microsoft’s own add-in library (office.js), which every Word add-in loads from Microsoft’s servers. Security teams generally accept this, since it comes from Microsoft. If a team insists on no outside loading at all, a copy of that library can be hosted internally, but Microsoft does not officially support that arrangement.",
      "The server needs a secure (https) address with a certificate that the organization’s computers trust. Organizations with their own IT department usually issue these routinely, so this requirement is normally simple in a managed corporate network.",
      "Word still has to allow the add-in, through a trusted shared folder or Microsoft 365 central deployment. The cybersecurity team would control that too.",
    ],
    "**What the cybersecurity team could do:**",
    [
      `Take the add-in’s files from the public repository at **${REPO}**: the **web** folder (12 small text files of HTML, JavaScript, CSS and settings, about 150 KB in all, plus 4 small icons) and **manifest.xml**. Everything is readable source code, with no compiled or hidden parts, so it can be reviewed line by line.`,
      "Host the **web** folder on an internal web server, reachable only inside the organization’s network. This could also be a web server on the meeting computer itself.",
      `Change the addresses in both copies of **manifest.xml** (one beside the **web** folder, the other inside it, which the add-in’s installation page offers for download) from **${SITE}** to the internal server. That is a simple find-and-replace.`,
      "Install that manifest through the organization’s usual process: the shared-folder method in section 12 of the User Manual, or Microsoft 365 central deployment.",
    ],
    "After that, nothing comes from Netlify or from the developer. The public repository can be checked periodically to see whether the add-in has been updated. The in-house copy changes only when the cybersecurity team puts new files on its internal server, so each update can be vetted and installed the same way.",
  ]],
  ["Reporting a security concern", [
    `A suspected security problem should not be posted as a GitHub Issue, where anyone could read it before it is fixed. Instead, on the **Security** tab of **${REPO}**, click **Report a vulnerability**. Only the developer sees the report, and is notified by email.`,
    "Other questions, problems and suggestions can be raised by opening an **Issue** on the project’s **Issues** tab (a free GitHub account is needed). Because **Issues** are public, do not include confidential information in them.",
  ]],
];

// ------------------------------------------------------------- SECURITY.md

function markdown() {
  const out = [`# ${TITLE}`, ""];
  for (const [heading, paras] of SECTIONS) {
    if (heading) out.push(`## ${heading}`, "");
    for (const p of paras) {
      if (Array.isArray(p)) out.push(...p.map((item) => `- ${item}`), "");
      else if (p.quote) out.push(`> ${p.quote}`, "");
      else out.push(p, "");
    }
  }
  return out.join("\n");
}

// ------------------------------------------------------------- Word copy

const FONT = "Calibri";

function runs(text, extra = {}) {
  return text.split(/(\*\*.+?\*\*)/).filter(Boolean).map((part) => {
    const b = /^\*\*(.+)\*\*$/.exec(part);
    return new TextRun({ text: b ? b[1] : part, bold: !!b || !!extra.bold, italics: !!extra.italics, font: FONT, size: extra.size || 22 });
  });
}

function word() {
  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: runs(TITLE, { bold: true, size: 32 }) }),
  ];
  for (const [heading, paras] of SECTIONS) {
    if (heading) children.push(new Paragraph({ keepNext: true, spacing: { before: 240, after: 120 }, children: runs(heading, { bold: true, size: 26 }) }));
    for (const p of paras) {
      if (Array.isArray(p)) {
        for (const item of p) children.push(new Paragraph({ bullet: { level: 0 }, spacing: { after: 80 }, children: runs(item) }));
      } else if (p.quote) {
        children.push(new Paragraph({ spacing: { after: 120 }, indent: { left: 720, right: 720 }, children: runs(p.quote, { italics: true }) }));
      } else {
        // A paragraph ending in a colon introduces what follows: keep them on one page.
        children.push(new Paragraph({ keepNext: /:\**$/.test(p), spacing: { after: 120 }, children: runs(p) }));
      }
    }
  }
  return new Document({
    creator: "Amendment Layers",
    title: TITLE,
    sections: [{
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ font: FONT, size: 20, children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES] })],
          })],
        }),
      },
      children,
    }],
  });
}

(async () => {
  const md = path.join(__dirname, "..", "SECURITY.md");
  fs.writeFileSync(md, markdown());
  console.log(`Wrote ${md}`);
  const docx = path.join(__dirname, `${TITLE}.docx`);
  fs.writeFileSync(docx, await Packer.toBuffer(word()));
  console.log(`Wrote ${docx}`);
})();
