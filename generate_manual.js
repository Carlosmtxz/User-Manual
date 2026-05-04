const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
  LevelFormat, ImageRun, TabStopType, LeaderType,
  Bookmark, InternalHyperlink
} = require('/home/claude/.npm-global/lib/node_modules/docx');
const fs = require('fs');

// ─── Color palette ────────────────────────────────────────────────────────────
const DARK_BLUE   = "1F3864";
const MID_BLUE    = "2E75B6";
const LIGHT_BLUE  = "D6E4F0";
const GRAY        = "595959";
const LIGHT_GRAY  = "F2F2F2";
const WHITE       = "FFFFFF";

// ─── Image helper ─────────────────────────────────────────────────────────────
const IMG_DIR = "/home/claude/User-Manual/FSDWB4_HMI_ScreenShots/";
const LOGO_PATH    = "/home/claude/User-Manual/logo solutions.png";
const MACHINE_PATH = "/home/claude/User-Manual/dualwb.jpg";
// Fit ~820x500 HMI images within 6.5" content width at 96 DPI
const IMG_W = 624;
const IMG_H = 380;
// Logo: 346x146px, display at 3 inches wide on cover (288px @ 96dpi)
const LOGO_W = 200;
const LOGO_H = 85;

const hmiImage = (filename, altTitle) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(`${IMG_DIR}${filename}`),
      transformation: { width: IMG_W, height: IMG_H },
      altText: { title: altTitle, description: altTitle, name: altTitle.replace(/\s+/g, '_') },
    })],
    spacing: { before: 40, after: 40 },
  });

const imgCaption = (text) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, italics: true, size: 20, font: "Arial", color: GRAY })],
    spacing: { before: 0, after: 60 },
  });

// ─── Helpers ──────────────────────────────────────────────────────────────────
const spacer = (before = 0, after = 0) =>
  new Paragraph({ children: [new TextRun("")], spacing: { before, after } });

const boldText = (text, size = 24, color = "000000") =>
  new TextRun({ text, bold: true, size, font: "Arial", color });

const normalText = (text, size = 22, color = "000000") =>
  new TextRun({ text, size, font: "Arial", color });

const placeholderRun = (text, size = 22) =>
  new TextRun({ text, size, font: "Arial", color: "7F7F7F", italics: true });

const sectionHeading = (text, bookmarkId) => {
  const run = new TextRun({ text, bold: true, size: 32, font: "Arial", color: WHITE });
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: bookmarkId
      ? [new Bookmark({ id: bookmarkId, children: [run] })]
      : [run],
    shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
    spacing: { before: 400, after: 200 },
    indent: { left: 180 },
  });
};

const subHeading = (text, bookmarkId, pageBreak = false) => {
  const run = new TextRun({ text, bold: true, size: 26, font: "Arial", color: DARK_BLUE });
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: bookmarkId
      ? [new Bookmark({ id: bookmarkId, children: [run] })]
      : [run],
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 2 } },
    pageBreakBefore: pageBreak,
  });
};

const bodyPara = (text, placeholder = false) =>
  new Paragraph({
    children: [placeholder ? placeholderRun(text) : normalText(text)],
    spacing: { before: 80, after: 120 },
  });

const numberedPara = (text, ref, placeholder = false) =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    children: [placeholder ? placeholderRun(text) : normalText(text)],
    spacing: { before: 60, after: 80 },
  });

const bulletPara = (text, ref, placeholder = false) =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    children: [placeholder ? placeholderRun(text) : normalText(text)],
    spacing: { before: 60, after: 80 },
  });

const infoBox = (label, text, fill = LIGHT_BLUE) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top:    { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE },
          left:   { style: BorderStyle.THICK,  size: 16, color: MID_BLUE },
          right:  { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" },
        },
        shading: { fill, type: ShadingType.CLEAR },
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 120, bottom: 120, left: 200, right: 120 },
        children: [
          new Paragraph({ children: [boldText(label + " ", 22, DARK_BLUE), normalText(text, 22, GRAY)] }),
        ],
      })],
    })],
  });

const kvTable = (rows) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3240, 6120],
    rows: rows.map(([label, value], i) =>
      new TableRow({
        children: [
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
            width: { size: 3240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 80 },
            children: [new Paragraph({ children: [boldText(label, 22)] })],
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
            width: { size: 6120, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 80 },
            children: [new Paragraph({ children: [normalText(value, 22)] })],
          }),
        ],
      })
    ),
  });

const stepCard = (num, title, description) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [960, 8400],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
          shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
          width: { size: 960, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(String(num), 32, WHITE)] })],
        }),
        new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
          shading: { fill: WHITE, type: ShadingType.CLEAR },
          width: { size: 8400, type: WidthType.DXA },
          margins: { top: 120, bottom: 120, left: 180, right: 120 },
          children: [
            new Paragraph({ children: [boldText(title, 24, DARK_BLUE)] }),
            new Paragraph({ children: [normalText(description, 22)], spacing: { before: 60 } }),
          ],
        }),
      ],
    })],
  });

const optionsTable = (rows) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3000, 6360],
    rows: [
      new TableRow({
        children: ["Option", "Description"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [3000, 6360][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...rows.map(([opt, desc, tags], i) =>
        new TableRow({
          children: [
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: 3000, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [
                new Paragraph({ children: [boldText(opt, 22, MID_BLUE)] }),
                ...(tags ? [new Paragraph({ children: [new TextRun({ text: tags, size: 18, font: "Arial", color: "7F7F7F", italics: true })] })] : []),
              ],
            }),
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: 6360, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [normalText(desc, 22)] })],
            }),
          ],
        })
      ),
    ],
  });

// ─── Cover page ───────────────────────────────────────────────────────────────
const logoExists = fs.existsSync(LOGO_PATH);
const coverPage = [
  spacer(1600),
  ...(logoExists ? [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(LOGO_PATH),
      transformation: { width: LOGO_W, height: LOGO_H },
      altText: { title: "Fox Solutions Logo", description: "Fox Solutions company logo", name: "FoxSolutionsLogo" },
    })],
    spacing: { after: 240 },
  })] : [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [boldText("Fox Solutions", 36, MID_BLUE)],
    spacing: { after: 240 },
  })]),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
        shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 200, bottom: 200, left: 300, right: 300 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText("FSDWB4 – Dual Wicketed Bagger", 32, WHITE)] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [normalText("User Manual", 24, WHITE)] }),
        ],
      })],
    })],
  }),
  spacer(40),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "jpg",
      data: fs.readFileSync(MACHINE_PATH),
      transformation: { width: 360, height: 291 },
      altText: { title: "FSDWB4 Dual Wicketed Bagger", description: "FSDWB4 machine photo", name: "FSDWB4_Machine" },
    })],
    spacing: { before: 0, after: 60 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [normalText("Version 1.0  |  Issue Date: 04-20-2026", 22, GRAY)],
    spacing: { after: 400 },
  }),
  kvTable([
    ["Machine Name:",   "Dual Wicketed Bagger"],
    ["Model Number:",   "FSDWB4"],
    ["Manufacturer:",   "Fox Solutions"],
  ]),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── Revision history ─────────────────────────────────────────────────────────
const revisionHistory = [
  sectionHeading("Revision History"),
  spacer(100),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1560, 1560, 3240, 3000],
    rows: [
      new TableRow({
        children: ["Rev.", "Date", "Description of Change", "Approved By"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [1560,1560,3240,3000][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...[
        ["1.0", "04-20-2026", "Creation of User Manual", "Carlos Martinez"],
      ].map(([rev, date, desc, approver], i) =>
        new TableRow({
          children: [rev, date, desc, approver].map((val, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [1560,1560,3240,3000][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [normalText(val, 22)] })],
            })
          ),
        })
      ),
    ],
  }),
  spacer(40),
  infoBox("ℹ PROGRAM COMPATIBILITY:", "This manual is intended for use with PLC program version FSDWB-P25B and its corresponding HMI version. Some features described in this document may be available in other program versions but are not guaranteed. If you are unsure which version is installed on your machine, refer to Setup → SYSTEM INFO on the HMI. Contact Fox Solutions technical support for version-specific guidance.", LIGHT_BLUE),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── Table of Contents ────────────────────────────────────────────────────────
const tocRow = (label, pg, indent = false, anchor = null) => {
  const textColor  = indent ? "444444" : DARK_BLUE;
  const fontSize   = indent ? 20 : 22;
  const isBold     = !indent;
  const labelRun   = new TextRun({ text: label,       font: "Arial", size: fontSize, bold: isBold, color: textColor });
  const pageRun    = new TextRun({ text: "\t" + pg,   font: "Arial", size: fontSize, bold: isBold, color: textColor });
  const children   = anchor
    ? [new InternalHyperlink({ anchor, children: [labelRun] }), pageRun]
    : [labelRun, pageRun];
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: 9360, leader: LeaderType.DOT }],
    spacing: { before: indent ? 60 : 140, after: 0 },
    indent: indent ? { left: 480 } : {},
    children,
  });
};

const tocSection = [
  sectionHeading("Table of Contents"),
  spacer(100),
  tocRow("1. Introduction",                                    "4",  false, "s1"),
  tocRow("1.1 Purpose of This Manual",                         "4",  true,  "s1-1"),
  tocRow("1.2 Machine Overview",                               "4",  true,  "s1-2"),
  tocRow("1.3 Intended Users",                                 "4",  true,  "s1-3"),
  tocRow("1.4 Technical Specifications",                       "4",  true,  "s1-4"),
  spacer(40),
  tocRow("2. Safety Overview",                                 "6",  false, "s2"),
  spacer(40),
  tocRow("3. Operating Instructions",                          "7",  false, "s3"),
  tocRow("3.1 Pre-Operation Checklist",                        "7",  true,  "s3-1"),
  tocRow("3.2 Starting the Machine",                           "7",  true,  "s3-2"),
  tocRow("3.3 During Normal Operation",                        "8",  true,  "s3-3"),
  tocRow("3.4 Stopping the Machine",                           "8",  true,  "s3-4"),
  tocRow("3.5 Emergency Stop Procedure",                       "9",  true,  "s3-5"),
  spacer(40),
  tocRow("4. HMI Screen Reference",                            "11", false, "s4"),
  tocRow("4.1 Main Screen",                                    "11", true,  "s4-1"),
  tocRow("4.2 Setup Menu – Page 1",                            "15", true,  "s4-2"),
  tocRow("4.2.1 I/O Control Panel",                            "16", true,  "s4-4"),
  tocRow("4.2.2 Test Parts Screen",                            "21", true,  "s4-9"),
  tocRow("4.2.3 Cleanup Screen",                               "23", true,  "s4-10"),
  tocRow("4.2.4 Options Menu",                                 "24", true,  "s4-11"),
  tocRow("4.2.5 Language Screen",                              "30", true,  "s4-2"),
  tocRow("4.2.6 Demo / Simulation Screen",                     "31", true,  "s4-2"),
  tocRow("4.2.7 System Info / Stats Screen",                   "33", true,  "s4-2"),
  tocRow("4.3 Setup Menu – Page 2",                            "34", true,  "s4-3"),
  tocRow("4.3.1 Sync Test Screen",                             "35", true,  "s4-3"),
  tocRow("4.3.2 Drive Status Screen",                          "36", true,  "s4-12"),
  tocRow("4.4 Parameters Screen",                              "37", true,  "s4-16"),
  tocRow("4.4.1 Hidden Timers Screen",                         "39", true,  "s4-16"),
  tocRow("4.4.2 Drive Parameters – WB1 Flat Belt",             "40", true,  "s4-12"),
  tocRow("4.4.3 Drive Parameters – WB1 V-Drive (Feed Belt)",   "41", true,  "s4-13"),
  tocRow("4.4.3.1 Pre-Alignment Setup Screen",                 "42", true,  "s4-13"),
  tocRow("4.4.4 Drive Parameters – WB2 Flat Belt",             "44", true,  "s4-14"),
  tocRow("4.4.5 Drive Parameters – WB2 V-Drive (Feed Belt)",   "45", true,  "s4-15"),
  tocRow("4.5 Recipes Screen",                                 "46", true,  "s4-17"),
  tocRow("4.6 Alarms Screen",                                  "47", true,  "s4-18"),
  spacer(40),
  tocRow("5. Sensor Setup & Configuration",                    "48", false, "s5"),
  tocRow("5.1 IFM OGD550 – Overview",                          "48", true,  "s5-1"),
  tocRow("5.2 Wiring Connections",                             "48", true,  "s5-2"),
  tocRow("5.3 Output Mode Configuration",                      "48", true,  "s5-3"),
  tocRow("5.4 Distance Setpoints",                             "49", true,  "s5-4"),
  tocRow("5.5 Verification",                                   "49", true,  "s5-5"),
  tocRow("5.6 Installer Notes",                                "49", true,  "s5-6"),
  tocRow("5.7 Pneumatic Sensor Assembly",                      "50", true,  "s5-7"),
  tocRow("5.8 SMC ZSE20B-T – Digital Vacuum Sensor",           "51", true,  "s5-8"),
  tocRow("5.9 SMC ISE20A-V – Digital Pressure Sensor",         "53", true,  "s5-9"),
  spacer(40),
  tocRow("6. Common Fault Codes & Troubleshooting",            "54", false, "s6"),
  spacer(40),
  tocRow("7. Machine Remote Connection",                       "57", false, "s7"),
  tocRow("7.1 Wi-Fi Setup",                                    "57", true,  "s7"),
  tocRow("7.2 Wired Ethernet Connection",                      "58", true,  "s7"),
  tocRow("7.3 VPN On/Off Switch",                              "58", true,  "s7"),
  tocRow("7.4 Network Requirements",                           "59", true,  "s7"),
  spacer(40),
  tocRow("8. Contact & Technical Support",                     "60", false, "s7"),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 1. Introduction ──────────────────────────────────────────────────────────
const introSection = [
  sectionHeading("1. Introduction", "s1"),
  spacer(100),
  subHeading("1.1 Purpose of This Manual", "s1-1"),
  bodyPara("This manual provides comprehensive operating instructions for the FSDWB4 Dual Wicketed Bagger. It is intended for use by qualified operators, technicians, and supervisors responsible for the day-to-day operation of the machine. Read all sections thoroughly before operating the equipment.", false),
  spacer(100),
  subHeading("1.2 Machine Overview", "s1-2"),
  bodyPara("The FSDWB4 Dual Wicketed Bagger is an automatic bagging machine designed for high-speed packaging of portioned products. The machine operates two bagging stations (WB1 and WB2), each connected to a single weigher that has multiple outlets. Each station sends a request signal to its corresponding weigher outlet when it is ready for product. Once the weigher delivers the correct portion, the station bags the product, seals the wicketed bag, and drops the finished bag onto an outfeed conveyor belt for downstream handling. For detailed performance specifications, refer to Section 1.4 – Technical Specifications."),
  spacer(40),
  infoBox("📋 MACHINE DOCUMENTATION:", "A QR code is located on the nameplate on the side of the control panel. Scanning it will give you access to the electrical schematics and user manual specific to your machine.", LIGHT_BLUE),
  spacer(40),
  subHeading("1.3 Intended Users", "s1-3"),
  bodyPara("This manual is intended for:"),
  bulletPara("Production operators responsible for running the FSDWB4 bagging machine during shifts", "bullets1"),
  bulletPara("Maintenance technicians performing routine upkeep or troubleshooting on the machine", "bullets1"),
  bulletPara("Production supervisors overseeing bagging line operations", "bullets1"),
  bulletPara("Quality assurance personnel monitoring bag weight and seal integrity", "bullets1"),
  spacer(100),
  subHeading("1.4 Technical Specifications", "s1-4"),
  spacer(40),
  kvTable([
    ["Machine Model:",        "FSDWB4"],
    ["Machine Type:",         "Dual Station Automatic Wicketed Bagger"],
    ["Number of Stations:",   "2 (WB1 and WB2)"],
    ["Rated Speed:",          "Up to 24 bags per minute per station (48 bags per minute combined)"],
    ["Portion Size:",         "1 lb – 10 lb (standard); larger portions available with special attachments — contact Fox Solutions for a custom solution"],
    ["Bag Type:",             "Wicketed bags"],
    ["Power Supply:",         "240 VAC, 20 A, 60 Hz"],
    ["Air Consumption:",      "18 CFM @ 90 PSI"],
    ["Operating Temp. Range:","40°F – 100°F (4°C – 38°C)"],
  ]),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 2. Safety Overview ───────────────────────────────────────────────────────
const safetyNote = [
  sectionHeading("2. Safety Overview", "s2"),
  spacer(100),
  infoBox("⚠ WARNING:", "Read and understand all safety precautions before operating this machine. Failure to comply may result in serious injury or death.", "FFF3CD"),
  spacer(40),
  bodyPara("Before operating this machine, all personnel must be familiar with the following safety requirements:"),
  bulletPara("Always wear appropriate Personal Protective Equipment (PPE): safety glasses, gloves, steel-toed boots, and hearing protection where required.", "bullets2"),
  bulletPara("Only trained and authorized personnel are permitted to operate this machine.", "bullets2"),
  bulletPara("Never bypass, defeat, or disable any safety guard, interlock, or emergency stop system.", "bullets2"),
  bulletPara("Lock Out / Tag Out (LOTO) procedures must be followed before performing any maintenance or clearing a jam.", "bullets2"),
  bulletPara("Keep the work area clean and free of obstructions at all times.", "bullets2"),
  bulletPara("Report any malfunction, unusual noise, or abnormal behavior to a supervisor immediately.", "bullets2"),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 3. HMI Screen Reference ──────────────────────────────────────────────────
const hmiSection = [ // eslint-disable-line no-unused-vars
  sectionHeading("4. HMI Screen Reference", "s4"),
  spacer(100),
  bodyPara("This section provides a visual reference for all screens available on the Human-Machine Interface (HMI) touchscreen. Screens are presented in the order they appear in the HMI navigation: Main, Setup (and submenus), Parameters (and submenus), Recipes, and Alarms."),
  spacer(40),

  // ── 4.1 Main Screen ──────────────────────────────────────────────────────────
  subHeading("4.1 Main Screen", "s4-1"),
  bodyPara("The Main Screen is the primary operating interface displayed during normal production. It provides real-time status and control for both bagging stations (WB1 and WB2). From this screen, operators can monitor the bagging cycle, enable or disable individual station subsystems, and start or stop production."),
  spacer(40),
  hmiImage("Main_Page.png", "Main Screen"),
  imgCaption("Figure 4.1a – Main Screen (Normal Operation)"),
  spacer(40),
  bodyPara("Key elements of the Main Screen:"),
  bulletPara("Navigation Bar (top): HOME, SETUP, PARAMETERS, RECIPES, ALARMS — provides access to all major screens. A cloud with a padlock icon may appear near the top of the navigation bar — this indicates an active remote connection to the HMI by Fox Solutions technical support for troubleshooting. No action is required from the operator when this icon is visible.", "bullets3"),
  bulletPara("START Button (top right): Green START initiates production.", "bullets3"),
  bulletPara("ESTOP Indicator (top right): The red ESTOP graphic is a visual representation of the physical E-stop button state — it is not a clickable HMI button. When the physical E-stop is pressed, this indicator will reflect that state and all machine motion will halt.", "bullets3"),
  bulletPara("Active Recipe Display: Shows the currently selected recipe (e.g., 1-2LB) and machine run status.", "bullets3"),
  bulletPara("WB1 / WB2 Panels: Each bagging station has independent controls including WEIGHER DUMP OFF and TEST DUMP buttons.", "bullets3"),
  bulletPara("KWIKLOK: Enables or disables the Kwiklok bag closure device for the respective bagging station.", "bullets3"),
  bulletPara("ENABLE WB: Activates or deactivates the bagging station from participating in the production cycle.", "bullets3"),
  bulletPara("AIR SUPPLY: Monitors or controls the pneumatic air supply status for each station.", "bullets3"),
  spacer(40),
    hmiImage("Main_Page_Estop.png", "Main Screen – E-Stop Activated"),
  imgCaption("Figure 4.1b – Main Screen (E-Stop Activated)"),
  spacer(40),
  bodyPara("When the physical E-stop is pressed, the ESTOP indicator on the Main Screen changes state to reflect that the machine has been halted. All motion stops immediately. The indicator will return to its normal state once the physical E-stop is released and the machine is reset."),
  spacer(40),
    hmiImage("Main_Page_ProductNotFound.png", "Main Screen – Product Not Found"),
  imgCaption("Figure 4.1c – Main Screen (Product Not Found)"),
  spacer(40),
  bodyPara("Triggered when a weigher dump signal is received but no product is detected on the feeding belts by the jam sensor or — where equipped — the V-conveyor or delivery sensors. Must be enabled via ENABLE PRODUCT DETECTION in the Options Menu (Section 4.2.5). When flagged, the affected station halts and presents the operator with two options:"),
  bulletPara("Orange button (Re-request) – Clears the saved dump signal and sends a new product request to the weigher. Use this if the feeding belts are empty and a fresh portion is needed.", "bullets3"),
  bulletPara("Blue button (Cycle Bag) – Assumes product was missed by the sensors and cycles the bag through the Kwiklok to move on to the next cycle. Use this if product is visibly present on the belts or in the bag.", "bullets3"),
  spacer(40),
  infoBox("⚠ WARNING:", "The operator must visually inspect the feeding belts and bags before selecting either option. Choosing incorrectly may result in an empty sealed bag or a missed portion reaching downstream.", "FFF3CD"),
  spacer(40),
    hmiImage("Main_Page_Weigher.png", "Main Screen – Weigher Status"),
  imgCaption("Figure 4.1d – Main Screen (Weigher Status)"),
  spacer(40),
  bodyPara("The Weigher Status view on the Main Screen displays the current communication state between the bagging station and its connected weigher outlet. When the station is actively requesting product, the weigher status indicator will reflect the pending request. Once the weigher confirms a dump, the station proceeds with the bagging cycle. Use this view to verify that the request/dump handshake is operating correctly during production."),
  spacer(40),

  // ── 4.2 Setup Menu – Page 1 ──────────────────────────────────────────────────
  subHeading("4.2 Setup Menu – Page 1", "s4-2", true),
  bodyPara("The Setup Menu is accessed by tapping the SETUP tab in the Navigation Bar. Page 1 contains diagnostic tools and system configuration options. Some functions require a technician login to access."),
  spacer(40),
  hmiImage("Setup_Page.png", "Setup Menu Page 1"),
  imgCaption("Figure 4.2 – Setup Menu, Page 1"),
  spacer(40),
  bodyPara("Available options on Setup Page 1:"),
  bulletPara("IO – View real-time input and output signal states; used for wiring verification and diagnostics (see Section 4.2.1).", "bullets4"),
  bulletPara("TEST PARTS – Manually actuate individual pneumatic and mechanical components for testing (see Section 4.2.2).", "bullets4"),
  bulletPara("SYSTEM INFO – Display CPU information and production counters (see Section 4.2.7).", "bullets4"),
  bulletPara("CLEAN UP – Initiate a cleanup or purge cycle to clear product from the machine (see Section 4.2.3).", "bullets4"),
  bulletPara("OPTIONS – Configure sensor bypass toggles and feature enable/disable settings (see Section 4.2.4).", "bullets4"),
  bulletPara("LANGUAGE – Change the HMI display language (see Section 4.2.5).", "bullets4"),
  bulletPara("SIMULATION – Run the machine through a full bagging cycle without live product or bags (see Section 4.2.6).", "bullets4"),
  spacer(40),

  // ── 4.2.1 I/O Control Panel ──────────────────────────────────────────────────
  subHeading("4.2.1 I/O Control Panel", "s4-4", true),
  bodyPara("The I/O Control Panel is accessed via Setup Menu → IO. It displays the real-time state of all machine inputs and outputs, allowing technicians to verify wiring, diagnose faults, and confirm signal integrity without requiring a PLC programmer."),
  spacer(40),
  hmiImage("Inputs1_Page.png", "I/O Inputs Page 1"),
  imgCaption("Figure 4.2.1a – I/O Control Panel – Inputs, Page 1"),
  spacer(40),
  bodyPara("Each row shows the signal name alongside a color-coded indicator: green = active (ON), grey = inactive (OFF). Use this screen to verify sensor states and confirm that E-stop, safety, and position switches are functioning correctly before starting production."),
  spacer(40),
    hmiImage("Inputs2_Page.png", "I/O Inputs Page 2"),
  imgCaption("Figure 4.2.1b – I/O Control Panel – Inputs, Page 2"),
  spacer(40),
  bodyPara("Inputs Page 2 continues the digital input signal display, covering additional sensors and switches. Review both input pages when troubleshooting unexpected machine behavior or verifying that all safety interlocks are responding correctly."),
  spacer(40),
    hmiImage("Outputs1_Page.png", "I/O Outputs Page 1"),
  imgCaption("Figure 4.2.1c – I/O Control Panel – Outputs, Page 1"),
  spacer(40),
  bodyPara("The Outputs pages display the real-time state of all digital output signals, including solenoid valves, motors, and other actuators controlled by the PLC. Use this screen to confirm whether the PLC is commanding a device to activate — helps isolate faults in control logic versus physical wiring."),
  spacer(40),
    hmiImage("Outputs2_Page.png", "I/O Pneumatic Outputs Page 2"),
  imgCaption("Figure 4.2.1d – I/O Pneumatic Panel – Outputs, Page 2"),
  spacer(40),
  bodyPara("The Pneumatic Panel Outputs display shows the state of pneumatic solenoid valves and related output signals. Use this page to verify that pneumatic actuators are being correctly commanded during the bagging cycle."),
  spacer(40),
    hmiImage("Outputs3_Page.png", "I/O Pneumatic Outputs Page 3"),
  imgCaption("Figure 4.2.1e – I/O Pneumatic Panel – Outputs, Page 3"),
  spacer(40),
  bodyPara("Page 3 continues the pneumatic output signals display for both WB1 and WB2 stations."),
  spacer(40),
  infoBox("ℹ NOTE:", "Output forcing (manually commanding an output from the HMI) may be available in certain modes. Only qualified technicians should use this function. Forcing outputs with the machine running can cause unexpected motion.", LIGHT_BLUE),
  spacer(40),

  // ── 4.2.2 Test Parts Screen ───────────────────────────────────────────────────
  subHeading("4.2.2 Test Parts Screen", "s4-9", true),
  bodyPara("The Test Parts Screen is accessed via Setup Menu → TEST PARTS. It allows technicians to manually activate individual pneumatic cylinders and mechanical components for each bagging station (WB1 and WB2) independently. Used during maintenance, commissioning, and troubleshooting to verify component operation without running a full production cycle."),
  spacer(40),
  hmiImage("TestParts_Page1.png", "Test Parts Screen"),
  imgCaption("Figure 4.2.2 – Test Parts Screen"),
  spacer(40),
  bodyPara("Each output on this screen can be triggered using one of two modes:"),
  bulletPara("LATCH – A single press activates the output and it remains ON. Press again to turn it OFF. Use this mode when you need to hold a component in position for inspection.", "bullets4"),
  bulletPara("PULSE – The output only stays ON while you press and hold the button (minimum 1 second). As soon as you release, the output turns back OFF. Use this for a controlled, momentary actuation.", "bullets4"),
  spacer(40),
  bodyPara("Continuous Cycle Testing (Yellow PULSE Button):"),
  bulletPara("A single yellow PULSE button is located on the right side panel, shared across all outputs. First press LATCH on the desired output, then press the yellow PULSE button to enable cycle mode. The actuator will automatically cycle on and off repeatedly — ON for the duration set in Time On, then OFF for the duration set in Time Off — until the PULSE button is disabled. Useful for verifying actuator timing, stroke consistency, and component endurance without manual intervention.", "bullets4"),
  spacer(40),
  infoBox("⚠ WARNING:", "Keep hands and body clear of all moving parts when using the Test Parts screen. Components will actuate immediately when triggered. Disable the yellow PULSE button and unlatch all outputs before performing any hands-on inspection.", "FFF3CD"),
  spacer(40),

  // ── 4.2.3 Cleanup Screen ──────────────────────────────────────────────────────
  subHeading("4.2.3 Cleanup Screen", "s4-10", true),
  bodyPara("The Cleanup Screen is accessed via Setup Menu → CLEAN UP. It allows operators to initiate a controlled cleanup or purge cycle to clear remaining product from the machine between production runs or during changeovers. Use this screen at the end of a shift or when switching between products."),
  spacer(40),
  hmiImage("CleanUp_Page.png", "Cleanup Screen"),
  imgCaption("Figure 4.2.3 – Cleanup Screen"),
  spacer(40),
  infoBox("ℹ NOTE:", "Always follow your facility's sanitation procedures after completing the cleanup cycle. The Cleanup Screen clears product from the machine path but does not replace a full sanitation wash-down where required.", LIGHT_BLUE),
  spacer(40),

  // ── 4.2.4 Options Menu ────────────────────────────────────────────────────────
  subHeading("4.2.4 Options Menu", "s4-11", true),
  bodyPara("The Options Menu is accessed via Setup Menu → OPTIONS (technician login required). It spans multiple pages and provides configuration settings for sensor bypasses, feature toggles, and machine behavior options. Changes here affect how the machine responds to sensor signals and which features are active during production."),
  spacer(40),
  bodyPara("Two symbols may appear alongside option names on this screen:"),
  bulletPara("Recipe Icon (folder symbol, shown to the left of the option name): The option value is saved per recipe. Turning it on or off applies only to the currently active recipe — different recipes can have different settings for this option.", "bullets3"),
  bulletPara("Chain Link Icon (shown between the WB1 and WB2 toggles): The option is shared between both bagging stations. Changing the setting on one station automatically applies the same value to the other.", "bullets3"),
  spacer(40),
  hmiImage("Options_1.png", "Options Menu Page 1"),
  imgCaption("Figure 4.2.4a – Options Menu, Page 1"),
  spacer(40),
  optionsTable([
    ["BUCKET JAM SENSOR BYPASS", "Disables the bucket jam detection sensor for the selected station. When ON, the machine will not stop if a jam is detected at the bucket. Use only during diagnostics or when the sensor has been confirmed faulty.", null],
    ["KWIKLOK ALARM BYPASS", "Bypasses the Kwiklok device alarm for the selected station. When ON, the machine continues running even if the Kwiklok reports a fault. Use with caution — this may mask a real mechanical issue.", null],
    ["COLLECTOR", "Enables the collector mechanism for the selected bagging station. When ON, the collector is active during the bagging cycle to assist with product handling. When the collector is enabled, the machine automatically disables Anticipation Time — because the collector holds a portion in close proximity to the feeding belts, requesting product early would cause the incoming portion to mix with the stored one, resulting in a double or mixed portion.", null],
    ["ENABLE LOAD BAGS BUTTON", "Activates a dedicated LOAD BAGS button on the Main Screen for the selected station. This allows operators to manually advance the bag wicket during setup or when reloading bags mid-production without starting a full cycle.", null],
  ]),
  spacer(40),
    hmiImage("Options_2.png", "Options Menu Page 2"),
  imgCaption("Figure 4.2.4b – Options Menu, Page 2"),
  spacer(40),
  optionsTable([
    ["ENABLE PRODUCT DETECTION", "Enables the Product Not Found detection feature. When ON, the machine monitors the feeding belts after each weigher dump signal using the jam sensor and — where equipped — the V-conveyor sensor and delivery sensor. If no product is detected within the expected time window after a dump, the machine triggers a Product Not Found alarm and halts the affected station. When OFF, the machine proceeds through the bagging cycle regardless of whether product was physically detected on the belts.", null],
    ["FEEDING BELTS EMPTY METHOD", "Determines how the machine detects that the feeding belts are empty between cycles. SENSOR uses a physical detection sensor to confirm the belts are clear; TIME uses a configurable timer to assume the belts are empty after a set duration has elapsed.", null],
    ["FEEDING BELTS STYLE", "Selects the type of feeding conveyor installed on the machine. V.CONVEYOR activates the V-Drive parameter screens; FLAT CONVEYOR activates the Flat Belt parameter screens. This selection determines which drive parameters are shown when accessing PARAMETERS → Next Page.", null],
    ["ENABLE BAG DROP CONTROL", "Enables the bag drop control feature per station. This option is used to prevent the machine from dropping a finished bag on top of a previously dropped bag still on the takeaway conveyor. When ON, the machine waits for confirmation that the conveyor is clear before releasing the next bag. This option should be paired with a photo eye sensor mounted on the takeaway conveyor to detect when the previous bag has cleared.", null],
  ]),
  spacer(40),
    hmiImage("Options_3.png", "Options Menu Page 3"),
  imgCaption("Figure 4.2.4c – Options Menu, Page 3"),
  spacer(40),
  optionsTable([
    ["KWIKLOK BELTS TIMEOUT", "When enabled, the bagger will automatically stop the Kwiklok belt if no bag has passed through it within a set period of time. This prevents the Kwiklok from running continuously when no product is present, reducing wear and alerting the operator to a potential feeding issue.", null],
    ["LARGE VOLUME", "Activates large volume mode for the bagging cycle. When enabled, the machine runs the feeding belts before requesting product from the weigher. This pre-motion helps streamline larger portions of product through the bagger and into the bag more smoothly, reducing jams and improving fill consistency for high-volume portions. Recipe-dependent — can be set differently per recipe. Linked — the same setting applies to both WB1 and WB2.", "Recipe | Linked"],
    ["ENABLE EXTERNAL HALT SIGNAL INPUT 15", "Enables the external halt signal on digital input 15. When ON, the machine will stop its cycle when this input receives a signal from an external device (e.g., a downstream conveyor or checkweigher). Linked — applies to both stations simultaneously.", "Linked"],
    ["RUN FEED BELTS BEFORE FINGER GRABS BAG", "When ON, the feed belts start running as soon as the bag opens, without waiting for the gripping finger to close and secure the bag. This reduces idle time and can increase machine throughput. However, if the bag is not seated tightly around the bucket, product falling into an unsecured bag may cause it to shift — potentially resulting in a spill or the finger failing to get a proper grip on the bag. Use this option only when bag presentation is consistent and reliable.", null],
  ]),
  spacer(40),
    hmiImage("Options_4.png", "Options Menu Page 4"),
  imgCaption("Figure 4.2.4d – Options Menu, Page 4"),
  spacer(40),
  optionsTable([
    ["PRE ALIGNMENT", "Activates the pre-alignment feature for V-Drive configurations only. When a portion is detected and present, the feed belts run briefly to align the product before it is fed into the bag. This creates a smoother, more controlled transition into the bagging cycle and reduces the risk of product arriving out of position. Additional pre-alignment timing settings can be found on the V-Drive Parameters page. Recipe-dependent — can be toggled per recipe. Linked — the same setting applies to both WB1 and WB2.", "Recipe | Linked"],
    ["BYPASS DELIVERY SENSOR", "Disables the delivery sensor that confirms product has cleared the feeding belts and entered the bucket. This bypass may be useful when handling products with excessive residual husks or debris that can trigger false jam alarms or cause delays. Use with caution — bypassing this sensor reduces the machine's ability to detect real jams.", null],
    ["LABEL DISPENSER", "Enables the label dispenser accessory. Only enable when the machine is equipped with this option. When ON, the machine drops a label or card into the bag from the top of the bucket at the appropriate point in the cycle. Recipe-dependent — can be enabled or disabled per recipe.", "Recipe"],
    ["KWIKLOK \"A\" MACHINE", "Configures the station to operate with a Kwiklok \"A\" style closure machine. Baggers are typically equipped with the newer \"B\" model — when using an \"A\" machine, the bagger cannot send a run signal as the A-model is not equipped to receive it. Additionally, instead of wiring the fault signal from a B machine, an external lock sensor must be installed into the Kwiklok fault input. Refer to the electrical schematics for wiring details. Linked — applies to both stations simultaneously.", "Linked"],
  ]),
  spacer(40),
    hmiImage("Options_5.png", "Options Menu Page 5"),
  imgCaption("Figure 4.2.4e – Options Menu, Page 5"),
  spacer(40),
  optionsTable([
    ["PUSHER", "Enables the pusher attachment, which assists the arm cylinder in pulling the bag into the Kwiklok. Designed for larger bag sizes — typically 10 lbs or more — where the standard arm cylinder may not provide sufficient force to seat the bag properly. The machine must be physically equipped with the pusher attachment before enabling this option. Recipe-dependent — can be enabled or disabled per recipe.", "Recipe"],
  ]),
  spacer(40),
  infoBox("⚠ WARNING:", "Only qualified technicians should modify Options settings. Incorrect configuration of sensor bypasses or feature toggles can cause unexpected machine behavior or compromise safety interlocks.", "FFF3CD"),
  spacer(40),

  // ── 4.2.5 Language Screen ─────────────────────────────────────────────────────
  subHeading("4.2.5 Language Screen", "s4-2", true),
  bodyPara("The Language Screen is accessed via Setup Menu → LANGUAGE. Select the desired display language and press the checkmark button to confirm. The HMI will immediately update all labels and text to the selected language."),
  spacer(40),
  hmiImage("Language_Page.png", "Language Screen"),
  imgCaption("Figure 4.2.5 – Language Screen"),
  spacer(40),
  infoBox("ℹ NOTE:", "English and Spanish are the available languages at this time.", LIGHT_BLUE),
  spacer(40),

  // ── 4.2.6 Demo / Simulation Screen ───────────────────────────────────────────
  subHeading("4.2.6 Demo / Simulation Screen", null, true),
  bodyPara("The Demo Screen is accessed via Setup Menu → SIMULATION. It provides two independent options that can be used separately or together to run the machine without live product or bags — useful for operator training, post-maintenance verification, or demonstrating machine operation safely. All cycle logic and mechanical motion execute normally."),
  spacer(40),
  hmiImage("Demo_Page.png", "Demo / Simulation Screen"),
  imgCaption("Figure 4.2.6 – Demo / Simulation Screen"),
  spacer(40),
  optionsTable([
    ["BYPASS WEIGHER\nOff / On", "When ON, the bagger does not wait for a live dump signal from the weigher to proceed. The machine still sends a physical request signal to the weigher — stop the weigher before enabling this option to prevent unwanted product from entering the bagger. A simulated dump signal is generated automatically after the time set in the Dump After Timer next to this option.", null],
    ["NO BAGS\nOff / On", "When ON, the bagger runs through the full cycle without bags loaded. When the vacuum generator activates and the vacuum cylinder extends, the machine simulates a Vacuum OK signal — as if a bag had been detected — after the time set in the Vacuum Sim Timer. This allows the cycle to complete normally without a bag present.", null],
  ]),
  spacer(40),
  infoBox("⚠ WARNING:", "Keep hands and body clear of all moving parts during simulation. All mechanical components actuate at normal operating speed. Safety interlocks and E-stop remain active at all times.", "FFF3CD"),
  spacer(40),
  infoBox("ℹ NOTE:", "Both BYPASS WEIGHER and NO BAGS must be turned OFF before returning to normal production. Leaving either option enabled will cause the machine to run without proper product or bag verification, resulting in empty or unsealed bags.", LIGHT_BLUE),
  spacer(40),

  // ── 4.2.7 System Info / Stats Screen ─────────────────────────────────────────
  subHeading("4.2.7 System Info / Stats Screen", null, true),
  bodyPara("The System Info Screen is accessed via Setup Menu → SYSTEM INFO. It displays basic CPU and firmware information alongside production bag counters for both bagging stations."),
  spacer(40),
  hmiImage("Stats_Page.png", "System Info / Stats Screen"),
  imgCaption("Figure 4.2.7 – System Info / Stats Screen"),
  spacer(40),
  bodyPara("The production counters are split into two categories:"),
  bulletPara("Current – Tracks bags completed since the machine was last powered on. Resets automatically on every power cycle.", "bullets3"),
  bulletPara("Permanent – Retains the total bag count across power cycles, providing a running lifetime total for each station.", "bullets3"),
  bulletPara("Total (bottom) – Displays the combined permanent count for both WB1 and WB2 stations.", "bullets3"),
  spacer(40),
  infoBox("ℹ NOTE:", "A RESET button will appear next to the Current counter when a technician security level is active. This allows the current count to be manually reset without affecting the permanent counter.", LIGHT_BLUE),
  spacer(40),

  // ── 4.3 Setup Menu – Page 2 ──────────────────────────────────────────────────
  subHeading("4.3 Setup Menu – Page 2", "s4-3", true),
  bodyPara("Setup Page 2 provides access to advanced diagnostic and drive configuration tools, primarily used by maintenance technicians during commissioning or troubleshooting."),
  spacer(40),
  hmiImage("Setup_Page2.png", "Setup Menu Page 2"),
  imgCaption("Figure 4.3 – Setup Menu, Page 2"),
  spacer(40),
  bodyPara("Available options on Setup Page 2:"),
  bulletPara("SYNC TEST – Tests the request/dump signal handshake between each bagging station and its weigher outlet. Use after initial setup or when troubleshooting communication issues (see Section 4.3.1).", "bullets5"),
  bulletPara("DRIVES – View the real-time status of all active drives (see Section 4.3.2).", "bullets5"),
  spacer(40),

  // ── 4.3.1 Sync Test Screen ───────────────────────────────────────────────────
  subHeading("4.3.1 Sync Test Screen", null, true),
  bodyPara("The Sync Test Screen is accessed via Setup Menu → SYNC TEST. It allows technicians to manually test and verify the request/dump signal handshake between each bagging station and its corresponding weigher outlet. Use this screen during initial commissioning or when troubleshooting communication issues between the bagger and the weigher."),
  spacer(40),
  hmiImage("Sync_Test_Page.png", "Sync Test Screen"),
  imgCaption("Figure 4.3.1 – Sync Test Screen"),
  spacer(40),
  bodyPara("Each station (WB1 and WB2) can be tested independently. The screen displays the current state of the request and dump signals in real time, allowing the technician to confirm that the weigher is receiving the request and responding with a valid dump signal. If the handshake fails or signals are not toggling as expected, check the weigher communication wiring and confirm the correct outlet is mapped to each station."),
  spacer(40),

  // ── 4.3.2 Drive Status Screen ────────────────────────────────────────────────
  subHeading("4.3.2 Drive Status Screen", "s4-12", true),
  bodyPara("The Drive Status Screen is accessed via Setup Menu → DRIVES. It provides a real-time overview of the health and communication status of all active drives on the machine. Only the drives relevant to the feed belt configuration currently selected in Options (Flat Belt or V-Drive) are displayed — drives not in use for the active configuration will not appear."),
  spacer(40),
  hmiImage("Drives_Page_Status.png", "Drive Status Screen"),
  imgCaption("Figure 4.3.2 – Drive Status Screen"),
  spacer(40),
  bodyPara("Use this screen to confirm that all active drives are communicating correctly and show no faults before starting production. A green status indicator means the drive is healthy and ready; a fault indicator means an error must be resolved before the machine can run. For detailed speed and acceleration settings for each drive, refer to Sections 4.4.3 through 4.4.6."),
  spacer(40),

  // ── 4.4 Parameters Screen ────────────────────────────────────────────────────
  subHeading("4.4 Parameters Screen", "s4-16", true),
  bodyPara("The Parameters Screen is accessed via the PARAMETERS tab in the Navigation Bar (technician login required). It displays and allows adjustment of timing values for both bagging stations (WB1 and WB2) that govern the product fall and feed cycle. When the COLLECTOR option is enabled, additional collector timing parameters appear on the screen. Tapping the Next Page button provides access to the drive parameters for WB1 and WB2, navigating automatically to Flat Belt or V-Drive screens depending on the drive style configured — no manual selection is required."),
  spacer(40),
  hmiImage("Parameters_Page1B.png", "Parameters Screen"),
  imgCaption("Figure 4.4 – Parameters Screen"),
  spacer(40),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3000, 6360],
    rows: [
      new TableRow({
        children: ["Parameter", "Description"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [3000, 6360][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...([
        ["WEIGHER FALL T", "Sets how long the machine waits after receiving a dump signal before expecting product on the feeding belts. This accounts for the travel time from the weigher outlet down to the belt. Adjust this value based on the weigher's sync type and physical layout — if the machine is getting dump signals too early or too late relative to product arrival, this is the first timer to tune. See weigher sync configuration for reference."],
        ["ANTICIPATION TIME", "Allows the machine to send a new product request to the weigher before the current feed belt cycle is fully complete, reducing idle time between cycles. This value is subtracted from the Feed Belt Run T — for example, if Feed Belt Run T is 100 and Anticipation Time is 10, the request is sent at 90. Use caution: setting this too high relative to Feed Belt Run T can cause the weigher to dump while product is still on the belts, resulting in a double dump. Note: when the COLLECTOR option is enabled, the machine automatically disables Anticipation Time to prevent mixed portions — see COLLECTOR in Options Menu (Section 4.2.4)."],
        ["FEED BELT RUN T", "Sets how long the feeding belts run to deliver product to the bag. Behavior depends on the Feeding Belts Empty Method set in Options: when set to SENSOR, this timer only starts once all feeding sensors (jam, delivery, and V-conveyor where equipped) confirm no product is present — ensuring the belts are clear before timing begins. When set to TIME, the timer starts as soon as the belts begin running, regardless of sensor state."],
        ["PRODUCT FALL TIME", "Starts when the feeding belts stop running. Gives the product time to fully clear the bucket and settle into the bag before the next stage of the cycle begins. Increase this value if product is arriving late in the bag or if the bucket is not fully clearing between cycles."],
      ].map(([param, desc], i) =>
        new TableRow({
          children: [param, desc].map((val, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [3000, 6360][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [j === 0 ? boldText(val, 22, MID_BLUE) : normalText(val, 22)] })],
            })
          ),
        })
      )),
    ],
  }),
  spacer(40),
  infoBox("ℹ NOTE – Timer Scale:", "All timer values are displayed in ms using a simplified scale where 100 ms = 1 second. For example, to set a 1-second delay enter 100, for 2 seconds enter 200, and for 1.5 seconds enter 150. This scale is used across all timing parameters for ease of programming.", LIGHT_BLUE),
  spacer(40),
  infoBox("ℹ NOTE:", "Collector timing parameters are only visible when the COLLECTOR option is enabled in Setup → OPTIONS. If the collector is not equipped or not enabled, these timers will not appear.", LIGHT_BLUE),
  spacer(40),

  subHeading("4.4.1 Hidden Timers Screen", null, true),
  bodyPara("The Hidden Timers Screen is accessed by pressing a hidden button in the bottom-right corner of the Parameters Screen. It contains core machine timing values that affect fundamental cycle behavior. These timers are factory-configured and do not require adjustment under normal operating conditions."),
  spacer(40),
  hmiImage("Hidden_Timers_Page.png", "Hidden Timers Screen"),
  imgCaption("Figure 4.4.2 – Hidden Timers Screen"),
  spacer(40),
  infoBox("⚠ WARNING:", "Do not modify any values on this screen without direct guidance from Fox Solutions technical support. Incorrect timer values can cause unpredictable machine behavior, missed cycles, or mechanical damage.", "FFF3CD"),
  spacer(40),

  subHeading("4.4.2 Drive Parameters – WB1 Flat Belt", null, true),
  bodyPara("Displayed automatically via PARAMETERS → Next Page when the Flat Belt drive style is configured for WB1. Allows adjustment of speed and acceleration for the WB1 flat belt conveyor motor."),
  spacer(40),
  hmiImage("WB1_FlatDrive_Parameters.png", "WB1 Flat Belt Drive Parameters"),
  imgCaption("Figure 4.4.3 – WB1 Flat Belt Drive Parameters"),
  spacer(40),
  bulletPara("SPEED (Hz) – Operating frequency of the flat belt drive motor. Increase to run the belt faster.", "bullets8"),
  bulletPara("ACCEL (s) – Ramp-up time from 0 Hz to the target speed. Lower values produce faster acceleration.", "bullets8"),
  bulletPara("DRIVE 1 STATUS – Indicates the health of the drive (green = OK, fault indicator if issue detected).", "bullets8"),
  bulletPara("TEST Button – Runs the drive momentarily at the configured speed for commissioning or verification.", "bullets8"),
  spacer(40),

  subHeading("4.4.3 Drive Parameters – WB1 V-Drive (Feed Belt)", null, true),
  bodyPara("Displayed automatically via PARAMETERS → Next Page when the V-Drive style is configured for WB1. The WB1 feed belt has two speed zones — an outer high-RPM zone and an inner low-RPM zone — each driven independently. When the PRE ALIGNMENT option is enabled in Options, a Pre-Alignment shortcut icon appears in the bottom-right corner of this screen."),
  spacer(40),
  hmiImage("WB1_VDrive_Parameters.png", "WB1 V-Drive Parameters"),
  imgCaption("Figure 4.4.3 – WB1 V-Drive Parameters (Feed Belt Outer / Inner)"),
  spacer(40),
  bulletPara("FEED BELT OUTER – HIGH RPM: Controls the outer section of the WB1 feed belt at high speed. Drive 1 status shown.", "bullets8"),
  bulletPara("FEED BELT INNER – LOW RPM: Controls the inner section of the WB1 feed belt at low speed. Drive 2 status shown.", "bullets8"),
  bulletPara("SPEED (Hz) and ACCEL (s): Configurable for each zone independently. These values apply to normal production feeding — not to the pre-alignment cycle.", "bullets8"),
  bulletPara("TEST Button: Runs each belt zone independently for verification.", "bullets8"),
  spacer(40),

  subHeading("4.4.3.1 Pre-Alignment Setup Screen", null, true),
  bodyPara("The Pre-Alignment Setup Screen is accessed by tapping the Pre-Alignment icon in the bottom-right corner of the V-Drive Parameters page. This feature is only available when the PRE ALIGNMENT option is enabled in the Options Menu (Section 4.2.4) and requires a V-Drive configuration. It runs the feeding belts in reverse before product is fed into the bag, improving product alignment and reducing the risk of jamming or bridging inside the bucket."),
  spacer(40),
  hmiImage("VDrive_Prealignment.png", "V-Drive Pre-Alignment Parameters"),
  imgCaption("Figure 4.4.3.1a – V-Drive Parameters with Pre-Alignment Icon"),
  spacer(40),
  hmiImage("VDrive_Prealignment2.png", "Pre-Alignment Setup Screen"),
  imgCaption("Figure 4.4.3.1b – Pre-Alignment Setup Screen"),
  spacer(40),
  bodyPara("Key settings on the Pre-Alignment Setup Screen:"),
  bulletPara("Belt Selection – Select which belt runs in reverse during the pre-alignment cycle. The selected belt is defined by the user.", "bullets8"),
  bulletPara("Cycles (1 or 2) – Select 1 or 2 pre-alignment cycles per portion. With 1 cycle, the selected belt runs in reverse for the duration set in Time 1. With 2 cycles, the first belt runs in reverse for Time 1, then the second belt runs in reverse for Time 2. Using 2 cycles adds time to the bagging sequence and will reduce machine output.", "bullets8"),
  bulletPara("Time 1 – Sets how long the first belt runs in reverse. Always visible.", "bullets8"),
  bulletPara("Time 2 – Sets how long the second belt runs in reverse. Only visible when 2 cycles is selected.", "bullets8"),
  bulletPara("Speed (Hz) and Accel (s) – These settings control belt speed and ramp-up time during the pre-alignment cycle only. They are independent from the normal production belt speeds configured on the V-Drive Parameters page and do not affect feeding speed during the bagging cycle.", "bullets8"),
  spacer(40),
  infoBox("ℹ NOTE:", "Using 2 cycles improves product positioning but reduces machine output. Use the minimum number of cycles needed for consistent alignment.", LIGHT_BLUE),
  spacer(40),

  subHeading("4.4.4 Drive Parameters – WB2 Flat Belt", null, true),
  bodyPara("Displayed automatically via PARAMETERS → Next Page when the Flat Belt drive style is configured for WB2. Mirrors the WB1 configuration on a separate drive channel."),
  spacer(40),
  hmiImage("WB2_FlatDrive_Parameters.png", "WB2 Flat Belt Drive Parameters"),
  imgCaption("Figure 4.4.5 – WB2 Flat Belt Drive Parameters"),
  spacer(40),
  bulletPara("SPEED (Hz) – Operating frequency of the WB2 flat belt drive.", "bullets9"),
  bulletPara("ACCEL (s) – Motor ramp-up time for the WB2 flat belt.", "bullets9"),
  bulletPara("DRIVE 4 STATUS – Health indicator for Drive 4 (WB2 flat belt).", "bullets9"),
  bulletPara("TEST Button – Activates the WB2 flat belt briefly for testing.", "bullets9"),
  spacer(40),

  subHeading("4.4.5 Drive Parameters – WB2 V-Drive (Feed Belt)", null, true),
  bodyPara("Displayed automatically via PARAMETERS → Next Page when the V-Drive style is configured for WB2. Like WB1, the WB2 feed belt has inner and outer speed zones controlled by separate drives (Drive 3 and Drive 4). The Pre-Alignment icon also appears on this screen when the PRE ALIGNMENT option is enabled — see Section 4.4.3.1 for setup details."),
  spacer(40),
  hmiImage("WB2_VDrive_Parameters.png", "WB2 V-Drive Parameters"),
  imgCaption("Figure 4.4.6 – WB2 V-Drive Parameters (Feed Belt Inner / Outer)"),
  spacer(40),
  bulletPara("FEED BELT INNER – LOW RPM: Controls the inner section of the WB2 feed belt at low speed. Drive 3 status shown.", "bullets9"),
  bulletPara("FEED BELT OUTER – HIGH RPM: Controls the outer section of the WB2 feed belt at high speed. Drive 4 status shown.", "bullets9"),
  bulletPara("SPEED (Hz) and ACCEL (s): Independently configurable for each belt zone.", "bullets9"),
  bulletPara("TEST Button: Runs each zone separately for commissioning verification.", "bullets9"),
  spacer(40),

  subHeading("4.5 Recipes Screen", "s4-17", true),
  bodyPara("The Recipes Screen is accessed from the RECIPES tab in the Navigation Bar. It allows operators to select the active weight recipe for the current production run. Each recipe corresponds to a product weight range and configures the machine timing and parameters accordingly."),
  spacer(40),
  hmiImage("Recipes_Page.png", "Recipes Screen"),
  imgCaption("Figure 4.5 – Recipes Screen"),
  spacer(40),
  bodyPara("Available recipe selections:"),
  bulletPara("1-2 LB — For portions weighing 1 to 2 pounds.", "bullets3"),
  bulletPara("3-4 LB — For portions weighing 3 to 4 pounds.", "bullets3"),
  bulletPara("5-6 LB — For portions weighing 5 to 6 pounds.", "bullets3"),
  bulletPara("7-8 LB — For portions weighing 7 to 8 pounds.", "bullets3"),
  bulletPara("9-10 LB — For portions weighing 9 to 10 pounds.", "bullets3"),
  bulletPara("11 LB+ — For portions weighing 11 pounds or more.", "bullets3"),
  spacer(40),
  infoBox("ℹ NOTE:", "Changing the active recipe during production will update machine timing parameters. Always verify output quality after a recipe change.", LIGHT_BLUE),
  spacer(40),
  infoBox("ℹ TIP – Renaming Recipes:", "Press and hold any recipe name on the Recipes Screen to rename it. Default names (e.g., 1-2 LB, 3-4 LB) can be customized to match your products or customer requirements.", LIGHT_BLUE),
  spacer(40),

  subHeading("4.6 Alarms Screen", "s4-18", false),
  bodyPara("The Alarms Screen is accessed from the ALARMS tab in the Navigation Bar. It displays a log of all active and historical alarms with their message and the time they were triggered. A red notification indicator appears next to the ALARMS tab in the Navigation Bar whenever there is an active alarm. Operators should review this screen when the machine stops unexpectedly or when the alarm indicator is visible."),
  spacer(40),
  hmiImage("Alarms_Page.png", "Alarms Screen"),
  imgCaption("Figure 4.6 – Alarms Screen"),
  spacer(40),
  bodyPara("Each alarm entry displays the following:"),
  bulletPara("Message – A plain-text description of the fault or warning condition.", "bullets3"),
  bulletPara("Activated – The date and time the alarm was triggered.", "bullets3"),
  spacer(40),
  bodyPara("Alarm entries are color-coded for quick identification:"),
  bulletPara("Red highlight – Alarm is currently active. The fault condition has not been resolved.", "bullets3"),
  bulletPara("Gray highlight – Alarm is inactive or has passed. The condition has cleared but the entry remains in the log.", "bullets3"),
  spacer(40),
  infoBox("⚠ WARNING:", "Do not clear alarms without first identifying and resolving the root cause. Repeatedly clearing an active alarm without fixing the fault may result in equipment damage or injury.", "FFF3CD"),
  spacer(40),
];

// ─── 4. Operating Instructions ────────────────────────────────────────────────
const operatingSection = [
  sectionHeading("3. Operating Instructions", "s3"),
  spacer(100),
  infoBox("ℹ BEFORE YOU BEGIN:", "Ensure you have completed the pre-operation checklist in Section 4.1 before starting the machine. Refer to Section 3 for descriptions of all HMI screens referenced below.", LIGHT_BLUE),
  spacer(40),

  subHeading("3.1 Pre-Operation Checklist", "s3-1"),
  bodyPara("Before each operating session, verify all of the following:"),
  spacer(40),
  ...[
    "Inspect machine for visible damage, loose bolts, or worn components. Report issues before proceeding.",
    "Confirm all safety guards and interlocks are in place and functioning correctly.",
    "Verify that the work area is clean and free of obstructions.",
    "Confirm the ESTOP button is functional — press and release to test.",
    "Ensure the correct recipe is selected on the HMI (see RECIPE display on the Main Screen).",
    "Verify DRIVE STATUS indicators on the Flat Belt and V Drive screens show no faults.",
  ].map(item => bulletPara(item, "bullets10")),
  spacer(40),

  subHeading("3.2 Starting the Machine", "s3-2"),
  bodyPara("Follow the steps below to start the machine safely:"),
  spacer(40),
  stepCard(1, "Power On", "Turn the main power switch to the ON position. The HMI touchscreen will illuminate and load the Main Screen."),
  spacer(40),
  stepCard(2, "Verify System Status", "Navigate to SETUP → SYSTEM INFO and confirm: Battery OK (green), I/O Configuration (green). If any indicator is red, contact maintenance before proceeding."),
  spacer(40),
  stepCard(3, "Select Recipe", "On the Main Screen, verify the RECIPE name shown in the status bar matches the product to be produced. The recipe controls portion timing and belt parameters for that product. Change via the RECIPES tab if needed."),
  spacer(40),
  stepCard(4, "Enable Bagging Stations", "On the Main Screen, toggle ENABLE WB and AIR SUPPLY to ON for WB1 and/or WB2. Each active station will begin sending REQUEST signals to its weigher outlet once production starts."),
  spacer(40),
  stepCard(5, "Start Production", "Press the green START button (top right of Main Screen). Each station will begin its cycle: requesting product from the weigher, receiving the portioned product, bagging and sealing it, then dropping the finished bag onto the outfeed conveyor. Each station is rated for up to 24 bags per minute (48 bags per minute combined across both stations). Monitor the first several cycles to confirm normal operation on both stations."),
  spacer(40),

  subHeading("3.3 During Normal Operation", "s3-3"),
  bodyPara("While the machine is running, monitor the Main Screen continuously for the following:"),
  bulletPara("WEIGHER DUMP OFF buttons — confirm each bagging station is cycling correctly.", "bullets11"),
  bulletPara("Dynamic Text product name — confirm the correct recipe remains active.", "bullets11"),
  bulletPara("Watch for any unexpected button state changes or alarm indicators on the Navigation Bar.", "bullets11"),
  bulletPara("Check physical bag output quality at regular intervals.", "bullets11"),
  spacer(40),
  infoBox("ℹ TIP:", "If output quality changes unexpectedly, navigate to PARAMETERS to verify timing values have not been altered.", LIGHT_BLUE),
  spacer(40),

  subHeading("3.4 Stopping the Machine", "s3-4"),
  bodyPara("At the end of each production run:"),
  numberedPara("Press the STOP / pause control to end the cycle after the current bag is completed.", "numbers1"),
  numberedPara("Allow all moving components to come to a full stop before approaching the machine.", "numbers1"),
  numberedPara("Record production totals from the SYSTEM INFO screen (Current WB1/WB2 and Total counts).", "numbers1"),
  numberedPara("Toggle ENABLE WB and AIR SUPPLY to OFF for both bays.", "numbers1"),
  numberedPara("Clean the machine and surrounding area.", "numbers1"),
  numberedPara("Turn the main power switch to the OFF position.", "numbers1"),
  numberedPara("Complete the operator production log.", "numbers1"),
  spacer(40),

  subHeading("3.5 Emergency Stop Procedure", "s3-5"),
  infoBox("🛑 EMERGENCY:", "In any emergency, press the physical red E-Stop button on the machine immediately. This cuts power to all motion systems. The ESTOP indicator on the Main Screen will reflect the activated state.", "FDDEDE"),
  spacer(40),
  bodyPara("After activating the ESTOP:"),
  numberedPara("Do NOT attempt to restart until the cause of the emergency has been identified and resolved.", "numbers2"),
  numberedPara("Alert your supervisor and/or the designated safety officer immediately.", "numbers2"),
  numberedPara("Complete an incident report before the machine is returned to service.", "numbers2"),
  numberedPara("To reset: twist the physical E-Stop button clockwise to release it, then follow the normal startup procedure from Section 3.2.", "numbers2"),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 5. Sensor Setup ──────────────────────────────────────────────────────────
const sensorSection = [
  sectionHeading("5. Sensor Setup & Configuration", "s5"),
  spacer(100),
  bodyPara("This section covers the setup and configuration of the sensors installed on the FSDWB4 machine. Each subsection identifies the specific sensor and provides wiring, output configuration, and setpoint instructions. Follow the appropriate subsection for the sensor being configured."),
  spacer(40),

  subHeading("5.1 IFM OGD550 – Overview", "s5-1"),
  bodyPara("The IFM OGD550 is an optical distance sensor with two configurable outputs. On this machine, it is configured to detect two distinct distances corresponding to the Jam Sensor and the Bucket Door Sensor. The sensor uses a 4-wire M12 connector."),
  spacer(40),
  infoBox("ℹ NOTE:", "This sensor is factory-configured for the FSDWB4 machine. Do not change the setpoints or output modes unless instructed by Fox Solutions technical support.", LIGHT_BLUE),
  spacer(40),

  subHeading("5.2 Wiring Connections (M12, 4-Wire)", "s5-2"),
  bodyPara("Connect the OGD550 using the M12 4-wire connector as follows:"),
  spacer(40),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 2500, 4860],
    rows: [
      new TableRow({
        children: ["Wire Color", "Signal", "Description"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [2000, 2500, 4860][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...([
        ["Brown",  "+24 VDC",             "Power Supply (+)"],
        ["Blue",   "0 VDC",              "Power Supply (–)"],
        ["Black",  "OUT1 – Jam Sensor",   "Output 1: active at 360 mm (Jam detection)"],
        ["White",  "OUT2 – Bucket Door",  "Output 2: active at 530 mm (Bucket door closed)"],
      ].map(([color, signal, desc], i) =>
        new TableRow({
          children: [color, signal, desc].map((val, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [2000, 2500, 4860][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [j === 0 ? boldText(val, 22, MID_BLUE) : normalText(val, 22)] })],
            })
          ),
        })
      )),
    ],
  }),
  spacer(40),
  infoBox("ℹ NOTE:", "Refer to the machine's electrical schematics for the actual terminal numbers at the control panel.", LIGHT_BLUE),
  spacer(40),

  subHeading("5.3 Output Mode Configuration", "s5-3"),
  bodyPara("Both outputs must be configured in PNP / Normally Open (NO) mode:"),
  spacer(40),
  bulletPara("OUT1 → PNP / NO  (Jam Sensor – Black wire)", "bullets5"),
  bulletPara("OUT2 → PNP / NO  (Bucket Door Sensor – White wire)", "bullets5"),
  spacer(40),
  infoBox("⚠ IMPORTANT:", "OUT2 (Bucket Door Sensor) should only be ON when the bucket door is fully closed. If OUT2 is active while the door is open, the sensor alignment must be rechecked.", "FFF3CD"),
  spacer(40),

  subHeading("5.4 Distance Setpoints", "s5-4"),
  bodyPara("The sensor is pre-configured with two distance setpoints that correspond to the two detection zones:"),
  spacer(40),
  bulletPara("SP1 (Output 1 – Jam Sensor): 360 mm — the Black wire output turns ON when an object is detected at this distance, indicating a potential product jam.", "bullets5"),
  bulletPara("SP2 (Output 2 – Bucket Door): 530 mm — the White wire output turns ON when the bucket door reaches the fully closed position.", "bullets5"),
  spacer(40),

  subHeading("5.5 Verification", "s5-5"),
  bodyPara("After wiring and configuring the sensor, verify correct operation using the following checks:"),
  spacer(40),
  bulletPara("At 360 mm: Black wire output (OUT1) = ON → Jam Sensor active.", "bullets5"),
  bulletPara("At 530 mm: White wire output (OUT2) = ON → Bucket Door Sensor active.", "bullets5"),
  bulletPara("Out of both detection ranges: Both outputs OFF.", "bullets5"),
  spacer(40),
  bodyPara("You can verify the sensor outputs in real time using the I/O Control Panel on the HMI (see Sections 4.4 through 4.8)."),
  spacer(40),

  subHeading("5.6 Installer Notes", "s5-6"),
  bulletPara("Clean the sensor lens before powering up. Dust or debris on the lens can cause false readings or missed detections.", "bullets5"),
  bulletPara("If readings are unstable or inconsistent, check the sensor's physical alignment and ensure there is no excessive ambient light interference in the detection zone.", "bullets5"),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),

  // ── Pneumatic Sensor Assembly ──────────────────────────────────────────────
  subHeading("5.7 Pneumatic Sensor Assembly", "s5-7"),
  bodyPara("The machine includes a stacked pneumatic sensor assembly mounted on a DIN rail bracket. The assembly contains four SMC digital sensors arranged from top to bottom in the following order:"),
  spacer(40),
  bulletPara("Position 1 (Top) – Pressure Sensor for WB1  |  SMC ISE20A-V  |  Schematic ref: PT1", "bullets5"),
  bulletPara("Position 2 – Vacuum Sensor for WB1  |  SMC ZSE20B-T  |  Schematic ref: PT2", "bullets5"),
  bulletPara("Position 3 – Pressure Sensor for WB2  |  SMC ISE20A-V  |  Schematic ref: PT3", "bullets5"),
  bulletPara("Position 4 (Bottom) – Vacuum Sensor for WB2  |  SMC ZSE20B-T  |  Schematic ref: PT4", "bullets5"),
  spacer(40),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(`${IMG_DIR}Pneumatic_Sensors_Assy.png`),
      transformation: { width: 300, height: 333 },
      altText: { title: "Pneumatic Sensor Assembly", description: "Pneumatic Sensor Assembly", name: "Pneumatic_Sensors_Assy" },
    })],
    spacing: { before: 120, after: 120 },
  }),
  imgCaption("Figure 5.7 – Pneumatic Sensor Assembly (Top to Bottom: WB1 Pressure, WB1 Vacuum, WB2 Pressure, WB2 Vacuum)"),
  spacer(40),
  bodyPara("Each sensor is wired independently to the machine's control panel. The assembly is located above the control panel. Refer to the electrical schematics for terminal assignments. Both sensor types share the same 5-wire color-coded wiring standard and 3-screen display interface, but monitor different pressure ranges."),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),

  // ── ZSE20B-T Vacuum Sensor ─────────────────────────────────────────────────
  subHeading("5.8 SMC ZSE20B-T – Digital Vacuum Sensor (WB1 & WB2)", "s5-8"),
  bodyPara("The SMC ZSE20B-T is a digital vacuum pressure switch used to monitor the vacuum level at WB1 and WB2. It features a 3-color, 3-screen LCD display and a PNP switch output for the configured setpoint."),
  spacer(40),

  subHeading("5.8.1 Wiring – ZSE20B-T"),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 2500, 4860],
    rows: [
      new TableRow({
        children: ["Wire Color", "Signal", "Description"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [2000, 2500, 4860][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...[
        ["Brown", "+24 VDC", "Power supply (+)"],
        ["Blue",  "0 VDC",   "Power supply (–)"],
        ["Black", "OUT1",    "Switch output – vacuum setpoint (SP1)"],
      ].map(([color, signal, desc], i) =>
        new TableRow({
          children: [color, signal, desc].map((v, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [2000, 2500, 4860][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [j === 0 ? boldText(v, 22, MID_BLUE) : normalText(v, 22)] })],
            })
          ),
        })
      ),
    ],
  }),
  spacer(40),

  subHeading("5.8.2 Setting SP1 – ZSE20B-T"),
  bodyPara("Use the front panel buttons to configure SP1. The preset for this machine is –50 kPa:"),
  spacer(40),
  bulletPara("Step 1 – Press and hold the SET button for 3 seconds to enter setting mode. The display will show 'SP1'.", "bullets5"),
  bulletPara("Step 2 – Use the UP (▲) and DOWN (▼) buttons to set SP1 to –50 kPa. Press SET to confirm.", "bullets5"),
  bulletPara("Step 3 – The sensor returns to run mode and displays the live vacuum reading.", "bullets5"),
  spacer(40),
  infoBox("ℹ NOTE:", "Display color changes based on output state: green = output ON (vacuum reached SP1), red = output OFF (no vacuum detected). The display unit can be changed by pressing UP and DOWN simultaneously while in run mode.", LIGHT_BLUE),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),

  // ── ISE20A-V Pressure Sensor ───────────────────────────────────────────────
  subHeading("5.9 SMC ISE20A-V – Digital Pressure Sensor (WB1 & WB2)", "s5-9"),
  bodyPara("The SMC ISE20A-V is a digital pressure switch used to monitor pneumatic pressure at both bagging stations (WB1 and WB2). Three units are installed on the assembly — one for WB1 and two for WB2. Each features a 3-color, 3-screen LCD display and a PNP switch output for the configured setpoint."),
  spacer(40),

  subHeading("5.9.1 Wiring – ISE20A-V"),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 2500, 4860],
    rows: [
      new TableRow({
        children: ["Wire Color", "Signal", "Description"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [2000, 2500, 4860][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...[
        ["Brown", "+24 VDC", "Power supply (+)"],
        ["Blue",  "0 VDC",   "Power supply (–)"],
        ["Black", "OUT1",    "Switch output – pressure setpoint (SP1)"],
      ].map(([color, signal, desc], i) =>
        new TableRow({
          children: [color, signal, desc].map((v, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [2000, 2500, 4860][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [j === 0 ? boldText(v, 22, MID_BLUE) : normalText(v, 22)] })],
            })
          ),
        })
      ),
    ],
  }),
  spacer(40),

  subHeading("5.9.2 Changing the Display Unit to PSI – ISE20A-V"),
  bodyPara("Before setting the setpoint, confirm the sensor is displaying in PSI. If the display shows MPa, kPa, or another unit, follow these steps to switch to PSI:"),
  spacer(40),
  bulletPara("Step 1 – While in run mode, press and hold both the UP (▲) and DOWN (▼) buttons simultaneously for 3 seconds until the unit selection screen appears.", "bullets5"),
  bulletPara("Step 2 – Use UP/DOWN to scroll through the available units: MPa → kPa → kgf/cm² → bar → psi.", "bullets5"),
  bulletPara("Step 3 – When 'psi' is shown, press SET to confirm. The display will return to run mode showing the live reading in PSI.", "bullets5"),
  spacer(40),

  subHeading("5.9.3 Setting SP1 – ISE20A-V"),
  bodyPara("SP1 is set to 75 PSI. Although the main air supply regulator is set to 90 PSI, the sensor threshold is intentionally set lower to account for the pressure drop that occurs when the vacuum generator activates — preventing false low-pressure faults during normal operation. Do not raise SP1 above 75 PSI."),
  spacer(40),
  bulletPara("Step 1 – Press and hold the SET button for 3 seconds to enter setting mode. The display will show 'SP1'.", "bullets5"),
  bulletPara("Step 2 – Use the UP (▲) and DOWN (▼) buttons to set SP1 to 75 PSI. Press SET to confirm.", "bullets5"),
  bulletPara("Step 3 – The sensor returns to run mode and displays the live pressure reading in PSI.", "bullets5"),
  spacer(40),
  infoBox("ℹ NOTE:", "Display color changes based on output state: green = pressure at or above SP1 (normal), red = pressure below SP1 (fault condition). Response time can be adjusted in the sensor's function settings to reduce chattering from brief pressure fluctuations during valve switching.", LIGHT_BLUE),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 6. Fault Codes ───────────────────────────────────────────────────────────
const faultSection = [
  sectionHeading("6. Common Fault Codes & Troubleshooting", "s6"),
  spacer(100),
  bodyPara("If an alarm is triggered, tap the ALARMS tab in the Navigation Bar to view the active alarm message. Use the table below to identify the fault and follow the recommended corrective actions. For issues not listed here, contact Fox Solutions technical support."),
  spacer(40),

  // Helper: fault card
  ...([
    {
      fault: "No Bag Detected",
      desc: "The vacuum generator activated but the vacuum sensor did not confirm a bag within the timeout period.",
      solutions: [
        "Adjust bag position — move the wicket plate up or down to ensure the suction cup makes full contact with the rear side of the bag.",
        "Adjust the vacuum cylinder assembly position — loosen the knobs and slide the assembly up or down, ensuring it clears the horseshoe on the bucket.",
        "Adjust the linear position of the vacuum tube — slide the tube forward so the suction cup reaches the bag properly.",
        "Vacuum timeout may be too short — access the Hidden Timers Screen (technician security required) to increase the vacuum timeout value.",
      ],
    },
    {
      fault: "No Air Pressure",
      desc: "The pressure sensor detected that pneumatic pressure has dropped below the configured setpoint.",
      solutions: [
        "Verify the main pneumatic shut-off valve is open and allowing airflow into the machine.",
        "Verify the main pressure regulator is set to the correct pressure (90 PSI).",
        "Verify the pressure sensor SP1 is set to the correct preset (75 PSI). See Section 5.9 for sensor configuration.",
      ],
    },
    {
      fault: "Kwiklok Fault",
      desc: "The Kwiklok bag closure machine has reported a fault condition.",
      solutions: [
        "If a Kwiklok B is equipped — refer to the fault chart engraved on top of the Kwiklok machine to identify and resolve the specific fault.",
        "If a Kwiklok A is equipped — verify the machine has sufficient locks loaded, or check the alignment of the external lock sensor.",
      ],
    },
    {
      fault: "Bucket Jammed",
      desc: "The bucket jam sensor detected an obstruction inside the bucket for longer than the allowed timeout.",
      solutions: [
        "Verify there is no product jammed inside the bucket and clear any obstruction.",
        "Verify the bucket jam sensor (IFM OGD550) is set to the correct setpoint. See Section 5 for sensor configuration.",
        "Jam timeout may need to be increased when running larger portions — access the Hidden Timers Screen (technician security required) to adjust.",
      ],
    },
    {
      fault: "V-Conv Blocked",
      desc: "The V-conveyor sensor detected product presence for longer than the allowed timeout, indicating a possible blockage.",
      solutions: [
        "Verify sensor alignment and integrity — confirm the V-conveyor sensor is properly aimed at its reflector.",
        "Verify the reflector mounted on the inside of the bucket is clean and correctly aligned with the V-conveyor sensor.",
      ],
    },
    {
      fault: "Delivery Blocked",
      desc: "The delivery sensor detected product presence for longer than the allowed timeout.",
      solutions: [
        "Verify sensor alignment — confirm the emitter and receiver are correctly aligned with each other.",
        "Delivery timeout may need to be increased when running larger portions — access the Hidden Timers Screen (technician security required) to adjust.",
      ],
    },
    {
      fault: "Unable to Reach Bag",
      desc: "The machine attempted to open a bag but could not confirm the bag was successfully grabbed.",
      solutions: [
        "Adjust bag position on the wicket — ensure bags are loaded correctly and the suction cup makes proper contact with the rear of the bag.",
        "Check vacuum system integrity — verify the vacuum generator is functioning and there are no air leaks in the vacuum line.",
        "Adjust the vacuum cylinder assembly position — loosen the knobs and slide the assembly up or down, ensuring it clears the horseshoe on the bucket.",
        "Adjust the linear position of the vacuum tube — slide the tube forward so the suction cup reaches the bag properly.",
      ],
      extras: [
        new Paragraph({ children: [boldText("Reference Diagrams:", 22, DARK_BLUE)], spacing: { before: 80, after: 40 } }),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [4680, 4680],
          rows: [new TableRow({
            children: [
              new TableCell({
                borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
                margins: { top: 40, bottom: 40, left: 40, right: 40 },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new ImageRun({ type: "png", data: fs.readFileSync(`${IMG_DIR}Vacuum CAD.png`), transformation: { width: 260, height: 200 }, altText: { title: "Vacuum Assembly CAD", description: "Vacuum Assembly CAD", name: "VacuumCAD" } })],
                  }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Vacuum Assembly", italics: true, size: 18, font: "Arial", color: GRAY })] }),
                ],
              }),
              new TableCell({
                borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
                margins: { top: 40, bottom: 40, left: 40, right: 40 },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new ImageRun({ type: "png", data: fs.readFileSync(`${IMG_DIR}BucketCAD.png`), transformation: { width: 260, height: 200 }, altText: { title: "Bucket CAD", description: "Bucket CAD", name: "BucketCAD" } })],
                  }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Bucket Assembly", italics: true, size: 18, font: "Arial", color: GRAY })] }),
                ],
              }),
            ],
          })],
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "ℹ NOTE: ", bold: true, size: 22, font: "Arial", color: DARK_BLUE }),
            new TextRun({ text: "After making adjustments, re-enable the station using the ENABLE WB toggle on the Main Screen to clear the alarm and resume the bagging cycle.", size: 22, font: "Arial" }),
          ],
          shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
          border: { left: { style: BorderStyle.THICK, size: 16, color: MID_BLUE } },
          indent: { left: 120 },
          spacing: { before: 80, after: 40 },
        }),
      ],
    },
    {
      fault: "Arm Cylinder Blocked",
      desc: "The arm cylinder was unable to reach or clear its position sensors (reed switches) within the expected time.",
      solutions: [
        "Release air pressure from the machine and manually move the arm cylinder to verify that both position sensors (extended and retracted) are reading correctly.",
        "Verify the finger assembly enters the side opening of the chute without contacting the side wall.",
      ],
    },
    {
      fault: "Vacuum Cylinder Blocked",
      desc: "The vacuum cylinder was unable to reach or clear its retracted position sensor (reed switch).",
      solutions: [
        "Release air pressure from the machine and manually move the vacuum cylinder to verify the retracted position sensor is reading correctly.",
      ],
    },
    {
      fault: "Bucket Position",
      desc: "The bucket door sensor (IFM OGD550 OUT2) is unable to confirm the bucket door is fully closed.",
      solutions: [
        "Verify the bucket sensor mounting bracket is not bent and the laser beam is reading the bucket door only when it is fully closed.",
        "Verify the bucket is clear of jams or obstructions that may be preventing the door from closing fully.",
      ],
    },
    {
      fault: "Label Dispenser — Out of Labels",
      desc: "The label dispenser vacuum was unable to pick a new label, indicating the label supply is empty or the vacuum system has an issue.",
      solutions: [
        "Verify the label supply is loaded and labels are positioned correctly for the dispenser to pick.",
        "Verify vacuum system integrity — check for leaks or blockages in the label dispenser vacuum line.",
      ],
    },
    {
      fault: "V-Conv Sensor Error",
      desc: "The V-conveyor sensor is reading intermittently — toggling on and off multiple times in rapid succession — indicating a possible alignment or contamination issue.",
      solutions: [
        "Verify V-conveyor sensor alignment with its reflector — adjust the sensor angle or position until a stable, consistent reading is achieved.",
        "Clean the sensor lens and reflector to remove dust or debris that may be causing intermittent readings.",
      ],
    },
  ].flatMap(({ fault, desc, solutions, extras }, i) => [
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [
        new TableRow({
          children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.THICK, size: 16, color: MID_BLUE }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
            shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 80, left: 200, right: 120 },
            children: [new Paragraph({ children: [boldText(fault, 24, DARK_BLUE)] })],
          })],
        }),
        new TableRow({
          children: [new TableCell({
            borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.THICK, size: 16, color: MID_BLUE }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
            shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 200, right: 120 },
            children: [new Paragraph({ children: [normalText(desc, 22, GRAY)] })],
          })],
        }),
        new TableRow({
          children: [new TableCell({
            borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: extras ? BorderStyle.NONE : BorderStyle.SINGLE, size: extras ? 0 : 4, color: extras ? "FFFFFF" : MID_BLUE }, left: { style: BorderStyle.THICK, size: 16, color: MID_BLUE }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
            shading: { fill: WHITE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: extras ? 40 : 120, left: 200, right: 120 },
            children: [
              new Paragraph({ children: [boldText("Possible Solutions:", 22, DARK_BLUE)], spacing: { before: 0, after: 60 } }),
              ...solutions.map(s => new Paragraph({
                numbering: { reference: "bullets3", level: 0 },
                children: [normalText(s, 22)],
                spacing: { before: 40, after: 60 },
              })),
            ],
          })],
        }),
        ...(extras ? [new TableRow({
          children: [new TableCell({
            borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE }, left: { style: BorderStyle.THICK, size: 16, color: MID_BLUE }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
            shading: { fill: WHITE, type: ShadingType.CLEAR },
            margins: { top: 40, bottom: 120, left: 200, right: 120 },
            children: extras,
          })],
        })] : []),
      ],
    }),
    spacer(40),
  ])),

  spacer(40),
];

// ─── 7. Remote Connection ─────────────────────────────────────────────────────
const remoteSection = [
  sectionHeading("7. Machine Remote Connection", "s7"),
  spacer(40),
  bodyPara("The FSDWB4 is equipped with an industrial VPN router (StrideLinx modem) mounted inside the control panel that enables secure remote access for diagnostics and technical support by Fox Solutions. The remote connection uses the StrideLinx cloud platform and can be established over Wi-Fi or a wired Ethernet connection."),
  spacer(40),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "jpeg",
      data: fs.readFileSync(`${IMG_DIR}Stridelinx_Modem.jpeg`),
      transformation: { width: 300, height: 260 },
      altText: { title: "StrideLinx Modem", description: "StrideLinx VPN Modem with USB Drive", name: "StridelinxModem" },
    })],
    spacing: { before: 40, after: 40 },
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Figure 7.0 – StrideLinx VPN Modem (USB drive shown on top)", italics: true, size: 20, font: "Arial", color: GRAY })], spacing: { before: 0, after: 40 } }),
  bodyPara("The modem has two indicator lights located above the ports:"),
  bulletPara("Wi-Fi Indicator (bars icon) — lights up purple when the modem has a good Wi-Fi connection to the customer's network.", "bullets3"),
  bulletPara("ACT Indicator — lights up purple when the modem is successfully reaching the StrideLinx server.", "bullets3"),
  spacer(40),

  subHeading("7.1 Wi-Fi Setup", null, false),
  bodyPara("To connect the machine to a Wi-Fi network, the customer must provide the following credentials to Fox Solutions:"),
  bulletPara("Network SSID (Wi-Fi network name)", "bullets3"),
  bulletPara("Wi-Fi password", "bullets3"),
  spacer(40),
  bodyPara("Fox Solutions will then generate a configuration file specific to the customer's network. To apply the configuration:"),
  numberedPara("Transfer the configuration file provided by Fox Solutions to the USB drive included with the modem.", "numbers3"),
  numberedPara("Power off the machine completely.", "numbers3"),
  numberedPara("Insert the USB drive into the modem inside the control panel.", "numbers3"),
  numberedPara("Power the machine back on — the modem will automatically apply the configuration on startup.", "numbers3"),
  spacer(40),
  infoBox("ℹ NOTE:", "We recommend completing the Wi-Fi setup during initial commissioning. However, the customer may configure it at any time with guidance from Fox Solutions technical support.", LIGHT_BLUE),
  spacer(40),
  infoBox("ℹ STRIDELINX USER ACCESS:", "Fox Solutions can add the customer as a user on the StrideLinx server, granting access to their machines only. This allows the customer to monitor their machine remotely at any time. Contact Fox Solutions to request user access.", LIGHT_BLUE),
  spacer(40),

  subHeading("7.2 Wired Ethernet Connection", null, false),
  bodyPara("If Wi-Fi is not preferred or available, the modem includes a wired Ethernet port (Port 1) for a direct hardwired connection to the customer's network. Connect a standard Ethernet cable from Port 1 on the modem to the customer's network switch or router."),
  spacer(40),

  subHeading("7.3 VPN On/Off Switch", null, false),
  bodyPara("The machine is equipped with an external VPN enable/disable switch mounted on the control panel. This switch allows the customer to control when the remote connection to the StrideLinx server is active."),
  spacer(40),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(`${IMG_DIR}External SW.png`),
      transformation: { width: 300, height: 240 },
      altText: { title: "External VPN Switch", description: "External VPN On/Off Switch", name: "ExternalSW" },
    })],
    spacing: { before: 40, after: 40 },
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Figure 7.3 – External VPN On/Off Switch", italics: true, size: 20, font: "Arial", color: GRAY })], spacing: { before: 0, after: 40 } }),
  bodyPara("We recommend keeping the VPN switch in the OFF position during normal production and enabling it only when remote assistance from Fox Solutions is required. This ensures the machine's network access remains fully under the customer's control."),
  spacer(40),
  infoBox("ℹ NOTE:", "The cloud with padlock icon on the HMI Main Screen will appear when the VPN switch is ON and an active remote session is in progress. See Section 4.1 for details.", LIGHT_BLUE),
  spacer(40),

  subHeading("7.4 Network Requirements", null, false),
  bodyPara("The StrideLinx VPN router establishes an outbound encrypted connection to the StrideLinx cloud service. No inbound firewall changes are required on the customer's network. The table below summarizes the network requirements:"),
  spacer(40),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3000, 6360],
    rows: [
      new TableRow({
        children: ["Category", "Requirement"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [3000, 6360][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...([
        ["Connection Type",    "Outbound connection initiated by machine VPN router"],
        ["Protocol",          "HTTPS / TLS encrypted VPN tunnel"],
        ["Port Required",     "TCP 443 (standard HTTPS)"],
        ["Firewall Changes",  "None required — no inbound port forwarding needed"],
        ["Authentication",    "User login required through StrideLinx portal"],
        ["Access Scope",      "Limited to machine control network (PLC, HMI, drives)"],
        ["Network Isolation", "Machine devices placed on isolated internal subnet"],
        ["Internet",          "Any standard internet connection with HTTPS access allowed"],
      ].map(([cat, req], i) =>
        new TableRow({
          children: [cat, req].map((val, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [3000, 6360][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [j === 0 ? boldText(val, 22, MID_BLUE) : normalText(val, 22)] })],
            })
          ),
        })
      )),
    ],
  }),
  spacer(40),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 8. Contact & Support ─────────────────────────────────────────────────────
const contactSection = [
  sectionHeading("8. Contact & Technical Support"),
  spacer(40),
  bodyPara("For technical assistance, spare parts, or service, please contact your authorized service representative or the manufacturer directly:"),
  spacer(40),
  kvTable([
    ["Manufacturer:",       "Fox Solutions"],
    ["Address:",            "2200 Fox Dr, McAllen, TX 78504"],
    ["Phone:",              "956-682-6176"],
    ["Email:",              "parts@foxbag.com  |  service@foxbag.com"],
    ["Website:",            "www.foxbag.com"],
    ["Service Hours:",      "Mon–Fri, 8:00 AM – 4:30 PM (Central Time)"],
  ]),
  spacer(40),
];

// ─── Build Document ────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets1",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets2",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets3",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets4",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets5",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets6",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets7",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets8",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets9",  levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets10", levels: [{ level: 0, format: LevelFormat.BULLET, text: "✓", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets11", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers1",  levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers2",  levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers3",  levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: WHITE },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: DARK_BLUE },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "Dual Wicketed Bagger (FSDWB4) – User Manual", font: "Arial", size: 18, color: GRAY }),
                new TextRun({ text: "\tFox Solutions", font: "Arial", size: 18, color: GRAY }),
              ],
              tabStops: [{ type: "right", position: 9360 }],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 4 } },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "Page ", font: "Arial", size: 18, color: GRAY }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: GRAY }),
                new TextRun({ text: " of ", font: "Arial", size: 18, color: GRAY }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: GRAY }),
              ],
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 4 } },
            }),
          ],
        }),
      },
      children: [
        ...coverPage,
        ...revisionHistory,
        ...tocSection,
        ...introSection,
        ...safetyNote,
        ...operatingSection,
        ...hmiSection,
        ...sensorSection,
        ...faultSection,
        ...remoteSection,
        ...contactSection,
      ],
    },
  ],
});

const outputPath = "/home/claude/User-Manual/FSDWB4-UM.docx";
const stream = Packer.toStream(doc);
const out = fs.createWriteStream(outputPath);
stream.pipe(out);
out.on("finish", () => console.log("✅ Document generated successfully."));
out.on("error", (err) => { console.error("❌ Error:", err); process.exit(1); });
