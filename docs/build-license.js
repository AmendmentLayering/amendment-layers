// Builds the add-in's license from the one text below:
//   ../LICENSE.md                                  (for the GitHub repository)
//   Amendment Layers Add-in License (draft).docx   (a Word copy for review)
// Edit the text here, then run:  node docs/build-license.js
//
// In the text, ***words*** are bold italic (as in the PolyForm original's
// No Liability section). A paragraph given as an array is a bulleted list.

const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require("docx");

const TITLE = "Amendment Layers Add-in License";

const SECTIONS = [
  ["Acceptance", [
    "In order to get any license under these terms, you must agree to them as both strict obligations and conditions to all your licenses.",
  ]],
  ["Copyright License", [
    "The licensor grants you a copyright license for the software to do everything you might do with the software that would otherwise infringe the licensor's copyright in it for any permitted purpose.  However, you may only distribute the software according to paragraph 3 (Distribution License) and make changes or new works based on the software according to paragraph 5 (Changes and New Works License).",
  ]],
  ["Distribution License", [
    "The licensor grants you an additional copyright license to distribute copies of the software.  Your license to distribute covers distributing the software with changes and new works permitted by paragraph 5 (Changes and New Works License).",
  ]],
  ["Notices", [
    "You must ensure that anyone who gets a copy of any part of the software from you also gets a copy of these terms.",
  ]],
  ["Changes and New Works License", [
    "The licensor grants you an additional copyright license to make changes and new works based on the software for any permitted purpose.",
  ]],
  ["Patent License", [
    "The licensor grants you a patent license for the software that covers patent claims the licensor can license, or becomes able to license, that you would infringe by using the software.",
  ]],
  ["Noncommercial Purposes", [
    "Any noncommercial purpose is a permitted purpose.",
  ]],
  ["Fair Use", [
    "You may have \"fair use\" rights for the software under the law. These terms do not limit them.",
  ]],
  ["No Other Rights", [
    "These terms do not allow you to sublicense or transfer any of your licenses to anyone else, or prevent the licensor from granting licenses to anyone else.  These terms do not imply any other licenses.",
  ]],
  ["Patent Defense", [
    "If you make any written claim that the software infringes or contributes to infringement of any patent, your patent license for the software granted under these terms ends immediately. If your company makes such a claim, your patent license ends immediately for work on behalf of your company.",
  ]],
  ["Violations", [
    "The first time you are notified in writing that you have violated any of these terms, or done anything with the software not covered by your licenses, your licenses can nonetheless continue if you come into full compliance with these terms, and take practical steps to correct past violations, within 32 days of receiving notice.  Otherwise, all your licenses end immediately.",
  ]],
  ["No Liability", [
    "***As far as the law allows, the software comes as is, without any warranty or condition, and the licensor will not be liable to you for any damages arising out of these terms or the use or nature of the software, under any kind of legal claim.***",
  ]],
  ["Definitions", [
    "The “licensor” is Thomas Balch Parliamentarian.",
    "A “commercial purpose” is:",
    "(a) selling, licensing, leasing, renting, or otherwise requiring payment of any fee or other consideration for the software, for any copy of it, for any change or new work based on it, or for access to any of them, including access through a hosted service; or",
    "(b) requiring payment of any fee or other consideration for a product or service that includes the software, or any change or new work based on it, as a component, where the payment is in any part for the software; however, inclusion with a paid product or service of a description of the software, or of a changed or new work based on the software, with information on how to access it, does not constitute its use for a commercial purpose as long as the software, or a changed or new work based on the software, is publicly available without charge regardless of whether anyone pays for the product or service that describes it and how to access it.",
    "A “noncommercial purpose” is any purpose that is not a commercial purpose. Use of the software in the course of providing services for which the user, or the user’s employer or principal, is paid is a noncommercial purpose, as long as no fee or other consideration is required for the software itself. Examples include:",
    [
      "presiding over or advising a meeting;",
      "serving as a parliamentarian, secretary, or recording officer;",
      "preparing, amending, or displaying documents under consideration; and",
      "otherwise conducting the business of any organization, whether nonprofit or for-profit.",
    ],
    "“Software” is the software the licensor makes available under these terms and includes the documentation accompanying the software:",
    [
      "README.md;",
      "Cybersecurity Vetting of the Amendment Layers Word Add-in (SECURITY.md);",
      "Amendment Layers – User Manual; and",
      "Amendment Layers – Quick Reference.",
    ],
    "“You” refers to the individual or entity agreeing to these terms.",
    "“Your company” is any legal entity, sole proprietorship, or other kind of organization that you work for, plus all organizations that have control over, are under the control of, or are under common control with that organization.  “Control” means ownership of substantially all the assets of an entity, or the power to direct its management and policies by vote, contract, or otherwise.  Control can be direct or indirect.",
    "“Your licenses” are all the licenses granted to you for the software under these terms.",
    "“Use” means anything you do with the software requiring one of your licenses.",
  ]],
];

// ------------------------------------------------------------- LICENSE.md

function markdown() {
  const out = [`# ${TITLE}`, ""];
  SECTIONS.forEach(([heading, paras], i) => {
    out.push(`## ${i + 1}. ${heading}`, "");
    for (const p of paras) {
      if (Array.isArray(p)) out.push(...p.map((item) => `- ${item}`), "");
      else out.push(p, "");
    }
  });
  return out.join("\n");
}

// ------------------------------------------------------------- Word copy

const FONT = "Calibri";

function runs(text, extra = {}) {
  return text.split(/(\*\*\*.+?\*\*\*)/).filter(Boolean).map((part) => {
    const bi = /^\*\*\*(.+)\*\*\*$/.exec(part);
    return new TextRun({ text: bi ? bi[1] : part, bold: !!bi || !!extra.bold, italics: !!bi, font: FONT, size: extra.size || 22 });
  });
}

function word() {
  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: runs(TITLE, { bold: true, size: 32 }) }),
  ];
  SECTIONS.forEach(([heading, paras], i) => {
    children.push(new Paragraph({ keepNext: true, spacing: { before: 240, after: 120 }, children: runs(`${i + 1}. ${heading}`, { bold: true, size: 24 }) }));
    for (const p of paras) {
      if (Array.isArray(p)) {
        for (const item of p) children.push(new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: runs(item) }));
      } else {
        const sub = /^\([a-z]\) /.test(p);
        children.push(new Paragraph({ spacing: { after: 120 }, indent: sub ? { left: 720 } : undefined, children: runs(p) }));
      }
    }
  });
  return new Document({
    creator: "Amendment Layers",
    title: TITLE,
    sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children }],
  });
}

(async () => {
  const md = path.join(__dirname, "..", "LICENSE.md");
  fs.writeFileSync(md, markdown());
  console.log(`Wrote ${md}`);
  const docx = path.join(__dirname, `${TITLE} (draft).docx`);
  fs.writeFileSync(docx, await Packer.toBuffer(word()));
  console.log(`Wrote ${docx}`);
})();
