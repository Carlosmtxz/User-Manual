const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
  LevelFormat, ImageRun, TabStopType, LeaderType,
  Bookmark, InternalHyperlink
} = require('/usr/local/lib/node_modules_global/lib/node_modules/docx');
const fs = require('fs');

// ─── Color palette ────────────────────────────────────────────────────────────
const DARK_BLUE   = "1F3864";
const MID_BLUE    = "2E75B6";
const LIGHT_BLUE  = "D6E4F0";
const GRAY        = "595959";
const LIGHT_GRAY  = "F2F2F2";
const WHITE       = "FFFFFF";

// ─── Image helper ─────────────────────────────────────────────────────────────
const IMG_DIR = "/sessions/keen-wonderful-bohr/mnt/User-Manual/FSDWB4-Manual/FSDWB4_HMI_ScreenShots/";
const LOGO_PATH    = "/sessions/keen-wonderful-bohr/mnt/User-Manual/FSDWB4-Manual/logo solutions.png";
const MACHINE_PATH = "/sessions/keen-wonderful-bohr/mnt/User-Manual/FSDWB4-Manual/dualwb.jpg";
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
    spacing: { before: 120, after: 120 },
  });

const imgCaption = (text) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, italics: true, size: 20, font: "Arial", color: GRAY })],
    spacing: { before: 0, after: 200 },
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

const subHeading = (text, bookmarkId) => {
  const run = new TextRun({ text, bold: true, size: 26, font: "Arial", color: DARK_BLUE });
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: bookmarkId
      ? [new Bookmark({ id: bookmarkId, children: [run] })]
      : [run],
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 2 } },
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
            children: [new Paragraph({ children: [placeholderRun(value, 22)] })],
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
            new Paragraph({ children: [placeholderRun(description, 22)], spacing: { before: 60 } }),
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
  spacer(200),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "jpg",
      data: fs.readFileSync(MACHINE_PATH),
      transformation: { width: 360, height: 291 },
      altText: { title: "FSDWB4 Dual Wicketed Bagger", description: "FSDWB4 machine photo", name: "FSDWB4_Machine" },
    })],
    spacing: { before: 0, after: 200 },
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
        ["1.1", "[YYYY-MM-DD]", "[Describe change]", "[Approver]"],
        ["1.2", "[YYYY-MM-DD]", "[Describe change]", "[Approver]"],
      ].map(([rev, date, desc, approver], i) =>
        new TableRow({
          children: [rev, date, desc, approver].map((val, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [1560,1560,3240,3000][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [placeholderRun(val, 22)] })],
            })
          ),
        })
      ),
    ],
  }),
  spacer(200),
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
  tocRow("1. Introduction",                              "4",  false, "s1"),
  tocRow("1.1 Purpose of This Manual",                   "4",  true,  "s1-1"),
  tocRow("1.2 Machine Overview",                         "4",  true,  "s1-2"),
  tocRow("1.3 Intended Users",                           "4",  true,  "s1-3"),
  tocRow("1.4 Technical Specifications",                 "4",  true,  "s1-4"),
  spacer(60),
  tocRow("2. Safety Overview",                           "6",  false, "s2"),
  spacer(60),
  tocRow("3. Operating Instructions",                    "7",  false, "s3"),
  tocRow("3.1 Pre-Operation Checklist",                  "7",  true,  "s3-1"),
  tocRow("3.2 Starting the Machine",                     "7",  true,  "s3-2"),
  tocRow("3.3 During Normal Operation",                  "8",  true,  "s3-3"),
  tocRow("3.4 Stopping the Machine",                     "8",  true,  "s3-4"),
  tocRow("3.5 Emergency Stop Procedure",                 "9",  true,  "s3-5"),
  spacer(60),
  tocRow("4. HMI Screen Reference",                      "10", false, "s4"),
  tocRow("4.1 Main Screen",                              "10", true,  "s4-1"),
  tocRow("4.2 Setup Menu – Page 1",                      "11", true,  "s4-2"),
  tocRow("4.3 Setup Menu – Page 2",                      "12", true,  "s4-3"),
  tocRow("4.4 I/O Control Panel – Inputs (Page 1)",      "13", true,  "s4-4"),
  tocRow("4.5 I/O Control Panel – Inputs (Page 2)",      "14", true,  "s4-5"),
  tocRow("4.6 I/O Control Panel – Outputs (Page 1)",     "15", true,  "s4-6"),
  tocRow("4.7 I/O Pneumatic Panel – Outputs (Page 2)",   "16", true,  "s4-7"),
  tocRow("4.8 I/O Pneumatic Panel – Outputs (Page 3)",   "17", true,  "s4-8"),
  tocRow("4.9 Test Parts Screen",                        "18", true,  "s4-9"),
  tocRow("4.10 Cleanup Screen",                          "20", true,  "s4-10"),
  tocRow("4.11 Options Menu",                            "20", true,  "s4-11"),
  tocRow("4.12 Drive Parameters – WB1 Flat Belt",        "25", true,  "s4-12"),
  tocRow("4.13 Drive Parameters – WB1 V-Drive (Feed Belt)", "26", true, "s4-13"),
  tocRow("4.14 Drive Parameters – WB2 Flat Belt",        "26", true,  "s4-14"),
  tocRow("4.15 Drive Parameters – WB2 V-Drive (Feed Belt)", "27", true, "s4-15"),
  tocRow("4.16 Parameters Screen",                       "28", true,  "s4-16"),
  tocRow("4.17 Recipes Screen",                          "30", true,  "s4-17"),
  tocRow("4.18 Alarms Screen",                           "31", true,  "s4-18"),
  spacer(60),
  tocRow("5. Sensor Setup & Configuration",              "33", false, "s5"),
  tocRow("5.1 IFM OGD550 – Overview",                    "33", true,  "s5-1"),
  tocRow("5.2 Wiring Connections",                       "33", true,  "s5-2"),
  tocRow("5.3 Output Mode Configuration",                "33", true,  "s5-3"),
  tocRow("5.4 Distance Setpoints",                       "34", true,  "s5-4"),
  tocRow("5.5 Verification",                             "34", true,  "s5-5"),
  tocRow("5.6 Installer Notes",                          "34", true,  "s5-6"),
  tocRow("5.7 Pneumatic Sensor Assembly",                "35", true,  "s5-7"),
  tocRow("5.8 SMC ZSE20B-T – Digital Vacuum Sensor",     "36", true,  "s5-8"),
  tocRow("5.9 SMC ISE20A-V – Digital Pressure Sensor",   "37", true,  "s5-9"),
  spacer(60),
  tocRow("6. Common Fault Codes & Troubleshooting",      "39", false, "s6"),
  spacer(60),
  tocRow("7. Contact & Technical Support",               "40", false, "s7"),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 1. Introduction ──────────────────────────────────────────────────────────
const introSection = [
  sectionHeading("1. Introduction", "s1"),
  spacer(100),
  subHeading("1.1 Purpose of This Manual", "s1-1"),
  bodyPara("This manual provides comprehensive operating instructions for the [Machine Name], Model [Model Number]. It is intended for use by qualified operators, technicians, and supervisors responsible for the day-to-day operation of the machine. Read all sections thoroughly before operating the equipment.", false),
  spacer(100),
  subHeading("1.2 Machine Overview", "s1-2"),
  bodyPara("The FSDWB4 Dual Wicketed Bagger is an automatic bagging machine designed for high-speed packaging of portioned products. The machine operates two bagging stations (WB1 and WB2), each connected to a single weigher that has multiple outlets. Each station sends a request signal to its corresponding weigher outlet when it is ready for product. Once the weigher delivers the correct portion, the station bags the product, seals the wicketed bag, and drops the finished bag onto an outfeed conveyor belt for downstream handling. For detailed performance specifications, refer to Section 1.4 – Technical Specifications."),
  spacer(100),
  subHeading("1.3 Intended Users", "s1-3"),
  bodyPara("This manual is intended for:"),
  bulletPara("Production operators responsible for running the FSDWB4 bagging machine during shifts", "bullets1"),
  bulletPara("Maintenance technicians performing routine upkeep or troubleshooting on the machine", "bullets1"),
  bulletPara("Production supervisors overseeing bagging line operations", "bullets1"),
  bulletPara("Quality assurance personnel monitoring bag weight and seal integrity", "bullets1"),
  spacer(100),
  subHeading("1.4 Technical Specifications", "s1-4"),
  spacer(80),
  kvTable([
    ["Machine Model:",        "FSDWB4"],
    ["Machine Type:",         "Dual Station Automatic Wicketed Bagger"],
    ["Number of Stations:",   "2 (WB1 and WB2)"],
    ["Rated Speed:",          "Up to 24 bags per minute per station (48 bags per minute combined)"],
    ["Portion Size:",         "1 lb – 10 lb (standard); larger portions available with special attachments — contact Fox Solutions for a custom solution"],
    ["Bag Type:",             "Wicketed bags"],
    ["Outfeed:",              "Conveyor belt"],
    ["Power Supply:",         "240 VAC, 20 A, 60 Hz"],
    ["Power Consumption:",    "[Enter power consumption]"],
    ["Machine Dimensions:",   "[Enter dimensions]"],
    ["Machine Weight:",       "[Enter weight]"],
    ["Operating Temp. Range:","[Enter operating temperature range]"],
  ]),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 2. Safety Overview ───────────────────────────────────────────────────────
const safetyNote = [
  sectionHeading("2. Safety Overview", "s2"),
  spacer(100),
  infoBox("⚠ WARNING:", "Read and understand all safety precautions before operating this machine. Failure to comply may result in serious injury or death.", "FFF3CD"),
  spacer(160),
  bodyPara("Before operating this machine, all personnel must be familiar with the following safety requirements:"),
  bulletPara("Always wear appropriate Personal Protective Equipment (PPE): safety glasses, gloves, steel-toed boots, and hearing protection where required.", "bullets2"),
  bulletPara("Only trained and authorized personnel are permitted to operate this machine.", "bullets2"),
  bulletPara("Never bypass, defeat, or disable any safety guard, interlock, or emergency stop system.", "bullets2"),
  bulletPara("Lock Out / Tag Out (LOTO) procedures must be followed before performing any maintenance or clearing a jam.", "bullets2"),
  bulletPara("Keep the work area clean and free of obstructions at all times.", "bullets2"),
  bulletPara("Report any malfunction, unusual noise, or abnormal behavior to a supervisor immediately.", "bullets2"),
  spacer(100),
  infoBox("ℹ NOTE:", "[Add any machine-specific safety notes here.]", LIGHT_BLUE),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 3. HMI Screen Reference ──────────────────────────────────────────────────
const hmiSection = [ // eslint-disable-line no-unused-vars
  sectionHeading("4. HMI Screen Reference", "s4"),
  spacer(100),
  bodyPara("This section provides a visual reference for all screens available on the Human-Machine Interface (HMI) touchscreen. Each screen is shown with a description of its purpose and key elements."),
  spacer(160),

  // 3.1 Main Screen
  subHeading("4.1 Main Screen", "s4-1"),
  bodyPara("The Main Screen is the primary operating interface displayed during normal production. It provides real-time status and control for both bagging stations (WB1 and WB2). From this screen, operators can monitor the bagging cycle, enable or disable individual station subsystems, and start or stop production."),
  spacer(120),
  hmiImage("Main_Page.png", "Main Screen"),
  imgCaption("Figure 4.1a – Main Screen (Normal Operation)"),
  spacer(80),
  bodyPara("Key elements of the Main Screen:"),
  bulletPara("Navigation Bar (top): HOME, SETUP, PARAMETERS, RECIPES, ALARMS — provides access to all major screens.", "bullets3"),
  bulletPara("START Button (top right): Green START initiates production.", "bullets3"),
  bulletPara("ESTOP Indicator (top right): The red ESTOP graphic is a visual representation of the physical E-stop button state — it is not a clickable HMI button. When the physical E-stop is pressed, this indicator will reflect that state and all machine motion will halt.", "bullets3"),
  bulletPara("Active Recipe Display: Shows the currently selected recipe (e.g., 1-2LB) and machine run status.", "bullets3"),
  bulletPara("WB1 / WB2 Panels: Each bagging station has independent controls including WEIGHER DUMP OFF and TEST DUMP buttons.", "bullets3"),
  bulletPara("KWIKLOK: Enables or disables the Kwiklok bag closure device for the respective bagging station.", "bullets3"),
  bulletPara("ENABLE WB: Activates or deactivates the bagging station from participating in the production cycle.", "bullets3"),
  bulletPara("AIR SUPPLY: Monitors or controls the pneumatic air supply status for each station.", "bullets3"),
  spacer(120),
  hmiImage("Main_Page_Estop.png", "Main Screen – E-Stop Activated"),
  imgCaption("Figure 4.1b – Main Screen (E-Stop Activated)"),
  spacer(80),
  bodyPara("When the physical E-stop is pressed, the ESTOP indicator on the Main Screen changes state to reflect that the machine has been halted. All motion stops immediately. The indicator will return to its normal state once the physical E-stop is released and the machine is reset."),
  spacer(200),

  // 3.2 Setup Menu – Page 1
  subHeading("4.2 Setup Menu – Page 1", "s4-2"),
  bodyPara("The Setup Menu is accessed by tapping the SETUP tab in the Navigation Bar. Page 1 contains diagnostic tools and system configuration options. Some functions require a technician login to access."),
  spacer(120),
  hmiImage("Setup_Page.png", "Setup Menu Page 1"),
  imgCaption("Figure 4.2 – Setup Menu, Page 1"),
  spacer(80),
  bodyPara("Available options on Setup Page 1:"),
  bulletPara("IO – View real-time input and output signal states; used for wiring verification and diagnostics (see Sections 4.4 through 4.8).", "bullets4"),
  bulletPara("TEST PARTS – Manually actuate individual pneumatic and mechanical components for testing (see Section 4.9).", "bullets4"),
  bulletPara("SYSTEM INFO – Display CPU program, firmware version, and production counters.", "bullets4"),
  bulletPara("CLEAN UP – Initiate a cleanup or purge cycle to clear product from the machine (see Section 4.10).", "bullets4"),
  bulletPara("OPTIONS – Configure sensor bypass toggles and feature enable/disable settings (see Section 4.11).", "bullets4"),
  bulletPara("LANGUAGE – Change the HMI display language.", "bullets4"),
  bulletPara("SIMULATION – Enable simulation mode for testing machine logic without live product or signals.", "bullets4"),
  spacer(200),

  // 3.3 Setup Menu – Page 2
  subHeading("4.3 Setup Menu – Page 2", "s4-3"),
  bodyPara("Setup Page 2 provides access to advanced diagnostic and drive configuration tools. These are primarily used by maintenance technicians during commissioning or troubleshooting."),
  spacer(120),
  hmiImage("Setup_Page2.png", "Setup Menu Page 2"),
  imgCaption("Figure 4.3 – Setup Menu, Page 2"),
  spacer(80),
  bodyPara("Available options on Setup Page 2:"),
  bulletPara("SYNC TEST – Tests the communication and synchronization between the bagging stations (WB1 and WB2) and the connected weigher. Use this after initial setup or when troubleshooting request/dump signal issues.", "bullets5"),
  spacer(200),

  // 3.4 I/O – Inputs Page 1
  subHeading("4.4 I/O Control Panel – Inputs (Page 1)", "s4-4"),
  bodyPara("The I/O Control Panel is accessed via Setup Menu → IO. It displays the real-time state of all machine inputs and outputs, allowing technicians to verify wiring, diagnose faults, and confirm signal integrity without requiring a PLC programmer. Inputs Page 1 covers the first set of digital input signals."),
  spacer(120),
  hmiImage("Inputs1_Page.png", "I/O Inputs Page 1"),
  imgCaption("Figure 4.4 – I/O Control Panel – Inputs, Page 1"),
  spacer(80),
  bodyPara("Each row in the panel shows the signal name alongside a color-coded indicator: green indicates the signal is active (ON), grey indicates inactive (OFF). Use this screen to quickly verify sensor states and confirm that E-stop, safety, and position switches are functioning correctly before starting production."),
  spacer(200),

  // 3.5 I/O – Inputs Page 2
  subHeading("4.5 I/O Control Panel – Inputs (Page 2)", "s4-5"),
  bodyPara("Inputs Page 2 continues the digital input signal display, covering additional sensors and switches not shown on Page 1."),
  spacer(120),
  hmiImage("Inputs2_Page.png", "I/O Inputs Page 2"),
  imgCaption("Figure 4.5 – I/O Control Panel – Inputs, Page 2"),
  spacer(80),
  bodyPara("Review both input pages when troubleshooting unexpected machine behavior or verifying that all safety interlocks are responding correctly."),
  spacer(200),

  // 3.6 I/O – Outputs Page 1
  subHeading("4.6 I/O Control Panel – Outputs (Page 1)", "s4-6"),
  bodyPara("The Outputs pages display the real-time state of all digital output signals, including solenoid valves, motors, and other actuators controlled by the PLC. Page 1 covers the first group of output signals in the I/O Control Panel."),
  spacer(120),
  hmiImage("Outputs1_Page.png", "I/O Outputs Page 1"),
  imgCaption("Figure 4.6 – I/O Control Panel – Outputs, Page 1"),
  spacer(80),
  bodyPara("Outputs are shown with their current state (ON/OFF). During troubleshooting, technicians can use this screen to confirm whether the PLC is commanding a device to activate, which helps isolate whether a fault is in the control logic or in the physical wiring and hardware."),
  spacer(200),

  // 3.7 I/O – Outputs Page 2 (Pneumatic Panel)
  subHeading("4.7 I/O Pneumatic Panel – Outputs (Page 2)", "s4-7"),
  bodyPara("The Pneumatic Panel Outputs display shows the state of the pneumatic solenoid valves and related output signals. This page is used to verify that pneumatic actuators are being correctly commanded during the bagging cycle."),
  spacer(120),
  hmiImage("Outputs2_Page.png", "I/O Pneumatic Outputs Page 2"),
  imgCaption("Figure 4.7 – I/O Pneumatic Panel – Outputs, Page 2"),
  spacer(200),

  // 3.8 I/O – Outputs Page 3 (Pneumatic Panel cont'd)
  subHeading("4.8 I/O Pneumatic Panel – Outputs (Page 3)", "s4-8"),
  bodyPara("Page 3 of the Pneumatic Panel Outputs continues the display of pneumatic output signals for both WB1 and WB2 stations."),
  spacer(120),
  hmiImage("Outputs3_Page.png", "I/O Pneumatic Outputs Page 3"),
  imgCaption("Figure 4.8 – I/O Pneumatic Panel – Outputs, Page 3"),
  spacer(80),
  infoBox("ℹ NOTE:", "Output forcing (manually commanding an output from the HMI) may be available in certain modes. Only qualified technicians should use this function. Forcing outputs with the machine running can cause unexpected motion.", LIGHT_BLUE),
  spacer(200),

  // 3.9 Test Parts Screen
  subHeading("4.9 Test Parts Screen", "s4-9"),
  bodyPara("The Test Parts Screen is accessed via Setup Menu → TEST PARTS. It allows technicians to manually activate individual pneumatic cylinders and mechanical components for each bagging station (WB1 and WB2) independently. This screen is used during maintenance, commissioning, and troubleshooting to verify component operation without running a full production cycle."),
  spacer(120),
  hmiImage("TestParts_Page1.png", "Test Parts Screen"),
  imgCaption("Figure 4.9 – Test Parts Screen"),
  spacer(80),
  bodyPara("Each output on this screen can be triggered using one of two modes:"),
  bulletPara("LATCH – A single press activates the output and it remains ON. Press again to turn it OFF. Use this mode when you need to hold a component in position for inspection.", "bullets4"),
  bulletPara("PULSE – The output only stays ON while you press and hold the button (minimum 1 second). As soon as you release, the output turns back OFF. Use this for a controlled, momentary actuation.", "bullets4"),
  spacer(80),
  bodyPara("Continuous Cycle Testing (Yellow PULSE Button):"),
  bulletPara("A single yellow PULSE button is located on the right side panel, shared across all outputs. This button enables automatic continuous cycling of whichever output is currently latched. To use it: first press LATCH on the desired output to activate and hold it ON, then press the yellow PULSE button to enable cycle mode. Once active, the actuator will automatically cycle on and off repeatedly — turning ON for the duration set in Time On, then turning OFF for the duration set in Time Off, and repeating until the PULSE button is disabled. This mode is useful for verifying actuator timing, stroke consistency, and component endurance without manual intervention.", "bullets4"),
  spacer(80),
  infoBox("⚠ WARNING:", "Keep hands and body clear of all moving parts when using the Test Parts screen. Components will actuate immediately when triggered. Disable the yellow PULSE button and unlatch all outputs before performing any hands-on inspection.", "FFF3CD"),
  spacer(200),

  // 3.10 Cleanup Screen
  subHeading("4.10 Cleanup Screen", "s4-10"),
  bodyPara("The Cleanup Screen is accessed via Setup Menu → CLEAN UP. It allows operators to initiate a controlled cleanup or purge cycle to clear remaining product from the machine between production runs or during changeovers. Use this screen at the end of a shift or when switching between products to ensure the machine is cleared before the next run."),
  spacer(120),
  hmiImage("CleanUp_Page.png", "Cleanup Screen"),
  imgCaption("Figure 4.10 – Cleanup Screen"),
  spacer(80),
  infoBox("ℹ NOTE:", "Always follow your facility's sanitation procedures after completing the cleanup cycle. The Cleanup Screen clears product from the machine path but does not replace a full sanitation wash-down where required.", LIGHT_BLUE),
  spacer(200),

  // 3.11 Options Menu
  subHeading("4.11 Options Menu", "s4-11"),
  bodyPara("The Options Menu is accessed via Setup Menu → OPTIONS (technician login required). It spans multiple pages and provides configuration settings for sensor bypasses, feature toggles, and machine behavior options. Changes here affect how the machine responds to sensor signals and which features are active during production."),
  spacer(80),
  bodyPara("Two symbols may appear alongside option names on this screen:"),
  bulletPara("Recipe Icon (folder symbol, shown to the left of the option name): The option value is saved per recipe. Turning it on or off applies only to the currently active recipe — different recipes can have different settings for this option.", "bullets3"),
  bulletPara("Chain Link Icon (shown between the WB1 and WB2 toggles): The option is shared between both bagging stations. Changing the setting on one station automatically applies the same value to the other.", "bullets3"),
  spacer(120),

  hmiImage("Options_1.png", "Options Menu Page 1"),
  imgCaption("Figure 4.11a – Options Menu, Page 1"),
  spacer(80),
  optionsTable([
    ["BUCKET JAM SENSOR BYPASS", "Disables the bucket jam detection sensor for the selected station. When ON, the machine will not stop if a jam is detected at the bucket. Use only during diagnostics or when the sensor has been confirmed faulty.", null],
    ["KWIKLOK ALARM BYPASS", "Bypasses the Kwiklok device alarm for the selected station. When ON, the machine continues running even if the Kwiklok reports a fault. Use with caution — this may mask a real mechanical issue.", null],
    ["COLLECTOR", "Enables the collector mechanism for the selected bagging station. When ON, the collector is active during the bagging cycle to assist with product handling.", null],
    ["ENABLE LOAD BAGS BUTTON", "Activates a dedicated LOAD BAGS button on the Main Screen for the selected station. This allows operators to manually advance the bag wicket during setup or when reloading bags mid-production without starting a full cycle.", null],
  ]),
  spacer(120),

  hmiImage("Options_2.png", "Options Menu Page 2"),
  imgCaption("Figure 4.11b – Options Menu, Page 2"),
  spacer(80),
  optionsTable([
    ["ENABLE PRODUCT DETECTION", "Activates the product detection sensor to verify that product is present in the bag before sealing. When OFF, the machine will seal the bag regardless of whether product was received.", null],
    ["FEEDING BELTS EMPTY METHOD", "Determines how the machine detects that the feeding belts are empty between cycles. SENSOR uses a physical detection sensor to confirm the belts are clear; TIME uses a configurable timer to assume the belts are empty after a set duration has elapsed.", null],
    ["FEEDING BELTS STYLE", "Selects the type of feeding conveyor installed on the machine. V.CONVEYOR activates the V-Drive parameter screens; FLAT CONVEYOR activates the Flat Belt parameter screens. This selection determines which drive parameters are shown when accessing PARAMETERS → Next Page.", null],
    ["ENABLE BAG DROP CONTROL", "Enables the bag drop control feature per station. This option is used to prevent the machine from dropping a finished bag on top of a previously dropped bag still on the takeaway conveyor. When ON, the machine waits for confirmation that the conveyor is clear before releasing the next bag. This option should be paired with a photo eye sensor mounted on the takeaway conveyor to detect when the previous bag has cleared.", null],
  ]),
  spacer(120),

  hmiImage("Options_3.png", "Options Menu Page 3"),
  imgCaption("Figure 4.11c – Options Menu, Page 3"),
  spacer(80),
  optionsTable([
    ["KWIKLOK BELTS TIMEOUT", "When enabled, the bagger will automatically stop the Kwiklok belt if no bag has passed through it within a set period of time. This prevents the Kwiklok from running continuously when no product is present, reducing wear and alerting the operator to a potential feeding issue.", null],
    ["LARGE VOLUME", "Activates large volume mode for the bagging cycle. When enabled, the machine runs the feeding belts before requesting product from the weigher. This pre-motion helps streamline larger portions of product through the bagger and into the bag more smoothly, reducing jams and improving fill consistency for high-volume portions. Recipe-dependent — can be set differently per recipe. Linked — the same setting applies to both WB1 and WB2.", "Recipe | Linked"],
    ["ENABLE EXTERNAL HALT SIGNAL INPUT 15", "Enables the external halt signal on digital input 15. When ON, the machine will stop its cycle when this input receives a signal from an external device (e.g., a downstream conveyor or checkweigher). Linked — applies to both stations simultaneously.", "Linked"],
    ["RUN FEED BELTS BEFORE FINGER GRABS BAG", "When ON, the feed belts start running as soon as the bag opens, without waiting for the gripping finger to close and secure the bag. This reduces idle time and can increase machine throughput. However, if the bag is not seated tightly around the bucket, product falling into an unsecured bag may cause it to shift — potentially resulting in a spill or the finger failing to get a proper grip on the bag. Use this option only when bag presentation is consistent and reliable.", null],
  ]),
  spacer(120),

  hmiImage("Options_4.png", "Options Menu Page 4"),
  imgCaption("Figure 4.11d – Options Menu, Page 4"),
  spacer(80),
  optionsTable([
    ["PRE ALIGNMENT", "Activates the pre-alignment feature for V-Drive configurations only. When a portion is detected and present, the feed belts run briefly to align the product before it is fed onto the belt. This creates a smoother, more controlled transition into the bagging cycle and reduces the risk of product arriving out of position. Additional pre-alignment timing settings can be found on the V-Drive Parameters page. Recipe-dependent — can be toggled per recipe. Linked — the same setting applies to both WB1 and WB2.", "Recipe | Linked"],
    ["BYPASS DELIVERY SENSOR", "Disables the delivery sensor that confirms a filled bag has cleared the station before the next cycle begins. Use only if the sensor is confirmed faulty — bypassing it may cause timing issues or double-fills if bags do not clear properly.", null],
    ["LABEL DISPENSER", "Enables the label dispenser accessory for the selected station. When ON, the machine triggers a label to be applied to each bag at the appropriate point in the cycle. Recipe-dependent — label dispensing can be enabled or disabled per recipe.", "Recipe"],
    ["KWIKLOK \"A\" MACHINE", "Configures the station to operate with a Kwiklok \"A\" style closure machine. This changes the Kwiklok actuation logic and timing to match the A-model unit. Linked — applies to both stations simultaneously.", "Linked"],
  ]),
  spacer(80),
  infoBox("⚠ WARNING:", "Only qualified technicians should modify Options settings. Incorrect configuration of sensor bypasses or feature toggles can cause unexpected machine behavior or compromise safety interlocks.", "FFF3CD"),
  spacer(200),

  // 3.12 WB1 Flat Belt Drive Parameters
  subHeading("4.12 Drive Parameters – WB1 Flat Belt", "s4-12"),
  bodyPara("The WB1 Flat Belt Drive Parameters screen is accessed via PARAMETERS → Next Page. This screen is displayed automatically when the Flat Belt drive style is configured for WB1. It displays and allows adjustment of the speed and acceleration settings for the WB1 flat belt conveyor motor. A TEST function is available to run the drive briefly for verification."),
  spacer(120),
  hmiImage("WB1_FlatDrive_Parameters.png", "WB1 Flat Belt Drive Parameters"),
  imgCaption("Figure 4.12 – WB1 Flat Belt Drive Parameters"),
  spacer(80),
  bulletPara("SPEED (Hz) – Operating frequency of the flat belt drive motor. Increase to run the belt faster.", "bullets8"),
  bulletPara("ACCEL (ms) – Ramp-up time from 0 Hz to the target speed. Lower values produce faster acceleration.", "bullets8"),
  bulletPara("DRIVE 1 STATUS – Indicates the health of the drive (green = OK, fault indicator if issue detected).", "bullets8"),
  bulletPara("TEST Button – Runs the drive momentarily at the configured speed for commissioning or verification.", "bullets8"),
  spacer(200),

  // 3.13 WB1 V-Drive Parameters
  subHeading("4.13 Drive Parameters – WB1 V-Drive (Feed Belt)", "s4-13"),
  bodyPara("The WB1 V-Drive Parameters screen is accessed via PARAMETERS → Next Page. This screen is displayed automatically when the V-Drive style is configured for WB1. It shows the configuration for the WB1 Feed Belt variable speed drives. The feed belt has two speed zones — an outer high-RPM zone and an inner low-RPM zone — each driven independently."),
  spacer(120),
  hmiImage("WB1_VDrive_Parameters.png", "WB1 V-Drive Parameters"),
  imgCaption("Figure 4.13 – WB1 V-Drive Parameters (Feed Belt Outer / Inner)"),
  spacer(80),
  bulletPara("FEED BELT OUTER – HIGH RPM: Controls the outer section of the WB1 feed belt at high speed. Drive 1 status shown.", "bullets8"),
  bulletPara("FEED BELT INNER – LOW RPM: Controls the inner section of the WB1 feed belt at low speed. Drive 2 status shown.", "bullets8"),
  bulletPara("SPEED (Hz) and ACCEL (ms): Configurable for each zone independently.", "bullets8"),
  bulletPara("TEST Button: Runs each belt zone independently for verification.", "bullets8"),
  spacer(200),

  // 3.14 WB2 Flat Belt Drive Parameters
  subHeading("4.14 Drive Parameters – WB2 Flat Belt", "s4-14"),
  bodyPara("The WB2 Flat Belt Drive Parameters screen is accessed via PARAMETERS → Next Page. This screen is displayed automatically when the Flat Belt drive style is configured for WB2. It displays speed and acceleration settings for the WB2 flat belt conveyor motor, mirroring the WB1 configuration on a separate drive channel."),
  spacer(120),
  hmiImage("WB2_FlatDrive_Parameters.png", "WB2 Flat Belt Drive Parameters"),
  imgCaption("Figure 4.14 – WB2 Flat Belt Drive Parameters"),
  spacer(80),
  bulletPara("SPEED (Hz) – Operating frequency of the WB2 flat belt drive.", "bullets9"),
  bulletPara("ACCEL (ms) – Motor ramp-up time for the WB2 flat belt.", "bullets9"),
  bulletPara("DRIVE 4 STATUS – Health indicator for Drive 4 (WB2 flat belt).", "bullets9"),
  bulletPara("TEST Button – Activates the WB2 flat belt briefly for testing.", "bullets9"),
  spacer(200),

  // 3.15 WB2 V-Drive Parameters
  subHeading("4.15 Drive Parameters – WB2 V-Drive (Feed Belt)", "s4-15"),
  bodyPara("The WB2 V-Drive Parameters screen is accessed via PARAMETERS → Next Page. This screen is displayed automatically when the V-Drive style is configured for WB2. Like WB1, the WB2 feed belt has inner and outer speed zones controlled by separate drives (Drive 3 and Drive 4)."),
  spacer(120),
  hmiImage("WB2_VDrive_Parameters.png", "WB2 V-Drive Parameters"),
  imgCaption("Figure 4.15 – WB2 V-Drive Parameters (Feed Belt Inner / Outer)"),
  spacer(80),
  bulletPara("FEED BELT INNER – LOW RPM: Controls the inner section of the WB2 feed belt at low speed. Drive 3 status shown.", "bullets9"),
  bulletPara("FEED BELT OUTER – HIGH RPM: Controls the outer section of the WB2 feed belt at high speed. Drive 4 status shown.", "bullets9"),
  bulletPara("SPEED (Hz) and ACCEL (ms): Independently configurable for each belt zone.", "bullets9"),
  bulletPara("TEST Button: Runs each zone separately for commissioning verification.", "bullets9"),
  spacer(200),

  // 3.16 Parameters Screen
  subHeading("4.16 Parameters Screen", "s4-16"),
  bodyPara("The Parameters Screen is accessed via the PARAMETERS tab in the Navigation Bar (technician login required). It displays and allows adjustment of timing values for both bagging stations (WB1 and WB2) that govern the product fall and feed cycle. Tapping the Next Page button on this screen provides access to the drive parameters for WB1 and WB2. The screen will automatically navigate to the Flat Belt or V-Drive parameters depending on the drive style configured for each station — no manual selection is required."),
  spacer(120),
  hmiImage("Parameters_Page.png", "Parameters Screen"),
  imgCaption("Figure 4.16 – Parameters Screen"),
  spacer(80),
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
        ["WEIGHER FALL T", "Time (ms) allowed for product to fall from the weigher head down into the bag."],
        ["ANTICIPATION TIME", "Pre-trigger time (ms) — how early the machine begins opening the bag before the weigher dump signal."],
        ["FEED BELT RUN T", "Duration (ms) the feed belt runs to deliver product to the weigher head."],
        ["PRODUCT FALL TIME", "Total time (ms) allowed for the product to fully clear the weigher and settle into the bag."],
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
  spacer(80),
  spacer(80),
  infoBox("ℹ NOTE – Timer Scale:", "All timer values are displayed in ms using a simplified scale where 100 ms = 1 second. For example, to set a 1-second delay enter 100, for 2 seconds enter 200, and for 1.5 seconds enter 150. This scale is used across all timing parameters for ease of programming.", LIGHT_BLUE),
  spacer(80),
  bodyPara("Each parameter is independently configurable for WB1 and WB2 to accommodate differences between the two stations."),
  spacer(200),

  // 3.17 Recipes Screen
  subHeading("4.17 Recipes Screen", "s4-17"),
  bodyPara("The Recipes Screen is accessed from the RECIPES tab in the Navigation Bar. It allows operators to select the active weight recipe for the current production run. Each recipe corresponds to a product weight range and configures the machine timing and parameters accordingly."),
  spacer(120),
  hmiImage("Recipes_Page.png", "Recipes Screen"),
  imgCaption("Figure 4.17 – Recipes Screen"),
  spacer(80),
  bodyPara("Available recipe selections:"),
  bulletPara("1-2 LB — For portions weighing 1 to 2 pounds.", "bullets3"),
  bulletPara("3-4 LB — For portions weighing 3 to 4 pounds.", "bullets3"),
  bulletPara("5-6 LB — For portions weighing 5 to 6 pounds.", "bullets3"),
  bulletPara("7-8 LB — For portions weighing 7 to 8 pounds.", "bullets3"),
  bulletPara("9-10 LB — For portions weighing 9 to 10 pounds.", "bullets3"),
  bulletPara("11 LB+ — For portions weighing 11 pounds or more.", "bullets3"),
  spacer(80),
  infoBox("ℹ NOTE:", "Changing the active recipe during production will update machine timing parameters. Always verify output quality after a recipe change.", LIGHT_BLUE),
  spacer(80),
  infoBox("ℹ TIP – Renaming Recipes:", "Recipe names can be customized to fit your production needs. To rename a recipe, press and hold the recipe name on the Recipes Screen — a text entry prompt will appear. The default names (e.g., 1-2 LB, 3-4 LB) are factory presets and can be edited to match your specific products or customer requirements.", LIGHT_BLUE),
  spacer(200),

  // 3.18 Alarms Screen
  subHeading("4.18 Alarms Screen", "s4-18"),
  bodyPara("The Alarms Screen is accessed from the ALARMS tab in the Navigation Bar. It displays a log of all active and historical alarms, including the alarm number, descriptive message, and activation time. Operators should review this screen when the machine stops unexpectedly or when an alarm indicator is visible."),
  spacer(120),
  hmiImage("Alarms_Page.png", "Alarms Screen"),
  imgCaption("Figure 4.18 – Alarms Screen"),
  spacer(80),
  bodyPara("Alarm log columns:"),
  bulletPara("Alarm No – A unique identifier for each alarm condition.", "bullets3"),
  bulletPara("Message – A plain-text description of the fault or warning condition.", "bullets3"),
  bulletPara("Activated – The date and time the alarm was triggered.", "bullets3"),
  spacer(80),
  infoBox("⚠ WARNING:", "Do not clear alarms without first identifying and resolving the root cause. Repeatedly clearing an active alarm without fixing the fault may result in equipment damage or injury.", "FFF3CD"),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 4. Operating Instructions ────────────────────────────────────────────────
const operatingSection = [
  sectionHeading("3. Operating Instructions", "s3"),
  spacer(100),
  infoBox("ℹ BEFORE YOU BEGIN:", "Ensure you have completed the pre-operation checklist in Section 4.1 before starting the machine. Refer to Section 3 for descriptions of all HMI screens referenced below.", LIGHT_BLUE),
  spacer(200),

  subHeading("3.1 Pre-Operation Checklist", "s3-1"),
  bodyPara("Before each operating session, verify all of the following:"),
  spacer(80),
  ...[
    "Inspect machine for visible damage, loose bolts, or worn components. Report issues before proceeding.",
    "Confirm all safety guards and interlocks are in place and functioning correctly.",
    "Check coolant, lubricant, and hydraulic fluid levels. Top up as required per specification.",
    "Verify that the work area is clean and free of obstructions.",
    "Confirm the ESTOP button is functional — press and release to test.",
    "Ensure the correct recipe is selected on the HMI (see RECIPE display on the Main Screen).",
    "Verify DRIVE STATUS indicators on the Flat Belt and V Drive screens show no faults.",
    "[Add machine-specific pre-start checks here]",
  ].map(item => bulletPara(item, "bullets10")),
  spacer(200),

  subHeading("3.2 Starting the Machine", "s3-2"),
  bodyPara("Follow the steps below to start the machine safely:"),
  spacer(120),
  stepCard(1, "Power On", "Turn the main power switch to the ON position. The HMI touchscreen will illuminate and load the Main Screen."),
  spacer(120),
  stepCard(2, "Verify System Status", "Navigate to SETUP → SYSTEM INFO and confirm: Battery OK (green), I/O Configuration (green). If any indicator is red, contact maintenance before proceeding."),
  spacer(120),
  stepCard(3, "Select Recipe", "On the Main Screen, verify the RECIPE name shown in the status bar matches the product to be produced. The recipe controls portion timing and belt parameters for that product. Change via the RECIPES tab if needed."),
  spacer(120),
  stepCard(4, "Enable Bagging Stations", "On the Main Screen, toggle ENABLE WB and AIR SUPPLY to ON for WB1 and/or WB2. Each active station will begin sending REQUEST signals to its weigher outlet once production starts."),
  spacer(120),
  stepCard(5, "Start Production", "Press the green START button (top right of Main Screen). Each station will begin its cycle: requesting product from the weigher, receiving the portioned product, bagging and sealing it, then dropping the finished bag onto the outfeed conveyor. Each station is rated for up to 24 bags per minute (48 bags per minute combined across both stations). Monitor the first several cycles to confirm normal operation on both stations."),
  spacer(200),

  subHeading("3.3 During Normal Operation", "s3-3"),
  bodyPara("While the machine is running, monitor the Main Screen continuously for the following:"),
  bulletPara("WEIGHER DUMP OFF buttons — confirm each bagging station is cycling correctly.", "bullets11"),
  bulletPara("Dynamic Text product name — confirm the correct recipe remains active.", "bullets11"),
  bulletPara("Watch for any unexpected button state changes or alarm indicators on the Navigation Bar.", "bullets11"),
  bulletPara("Check physical bag output quality at regular intervals.", "bullets11"),
  spacer(120),
  infoBox("ℹ TIP:", "If output quality changes unexpectedly, navigate to PARAMETERS to verify timing values have not been altered.", LIGHT_BLUE),
  spacer(200),

  subHeading("3.4 Stopping the Machine", "s3-4"),
  bodyPara("At the end of each production run:"),
  numberedPara("Press the STOP / pause control to end the cycle after the current bag is completed.", "numbers1"),
  numberedPara("Allow all moving components to come to a full stop before approaching the machine.", "numbers1"),
  numberedPara("Record production totals from the SYSTEM INFO screen (Current WB1/WB2 and Total counts).", "numbers1"),
  numberedPara("Toggle ENABLE WB and AIR SUPPLY to OFF for both bays.", "numbers1"),
  numberedPara("Clean the machine and surrounding area.", "numbers1"),
  numberedPara("Turn the main power switch to the OFF position.", "numbers1"),
  numberedPara("Complete the operator production log.", "numbers1"),
  spacer(200),

  subHeading("3.5 Emergency Stop Procedure", "s3-5"),
  infoBox("🛑 EMERGENCY:", "In any emergency, press the physical red E-Stop button on the machine immediately. This cuts power to all motion systems. The ESTOP indicator on the Main Screen will reflect the activated state.", "FDDEDE"),
  spacer(160),
  bodyPara("After activating the ESTOP:"),
  numberedPara("Do NOT attempt to restart until the cause of the emergency has been identified and resolved.", "numbers2"),
  numberedPara("Alert your supervisor and/or the designated safety officer immediately.", "numbers2"),
  numberedPara("If there is an injury, call [Emergency Number / First Aid contact] immediately.", "numbers2"),
  numberedPara("Complete an incident report before the machine is returned to service.", "numbers2"),
  numberedPara("To reset: twist the physical E-Stop button clockwise to release it, then follow the normal startup procedure from Section 3.2.", "numbers2"),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 5. Sensor Setup ──────────────────────────────────────────────────────────
const sensorSection = [
  sectionHeading("5. Sensor Setup & Configuration", "s5"),
  spacer(100),
  bodyPara("This section covers the setup and configuration of the sensors installed on the FSDWB4 machine. Each subsection identifies the specific sensor and provides wiring, output configuration, and setpoint instructions. Follow the appropriate subsection for the sensor being configured."),
  spacer(160),

  subHeading("5.1 IFM OGD550 – Overview", "s5-1"),
  bodyPara("The IFM OGD550 is an optical distance sensor with two configurable outputs. On this machine, it is configured to detect two distinct distances corresponding to the Jam Sensor and the Bucket Door Sensor. The sensor uses a 4-wire M12 connector."),
  spacer(80),
  infoBox("ℹ NOTE:", "This sensor is factory-configured for the FSDWB4 machine. Do not change the setpoints or output modes unless instructed by Fox Solutions technical support.", LIGHT_BLUE),
  spacer(160),

  subHeading("5.2 Wiring Connections (M12, 4-Wire)", "s5-2"),
  bodyPara("Connect the OGD550 using the M12 4-wire connector as follows:"),
  spacer(80),
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
  spacer(80),
  infoBox("ℹ NOTE:", "Refer to the machine's electrical schematics for the actual terminal numbers at the control panel.", LIGHT_BLUE),
  spacer(160),

  subHeading("5.3 Output Mode Configuration", "s5-3"),
  bodyPara("Both outputs must be configured in PNP / Normally Open (NO) mode:"),
  spacer(80),
  bulletPara("OUT1 → PNP / NO  (Jam Sensor – Black wire)", "bullets5"),
  bulletPara("OUT2 → PNP / NO  (Bucket Door Sensor – White wire)", "bullets5"),
  spacer(80),
  infoBox("⚠ IMPORTANT:", "OUT2 (Bucket Door Sensor) should only be ON when the bucket door is fully closed. If OUT2 is active while the door is open, the sensor alignment must be rechecked.", "FFF3CD"),
  spacer(160),

  subHeading("5.4 Distance Setpoints", "s5-4"),
  bodyPara("The sensor is pre-configured with two distance setpoints that correspond to the two detection zones:"),
  spacer(80),
  bulletPara("SP1 (Output 1 – Jam Sensor): 360 mm — the Black wire output turns ON when an object is detected at this distance, indicating a potential product jam.", "bullets5"),
  bulletPara("SP2 (Output 2 – Bucket Door): 530 mm — the White wire output turns ON when the bucket door reaches the fully closed position.", "bullets5"),
  spacer(160),

  subHeading("5.5 Verification", "s5-5"),
  bodyPara("After wiring and configuring the sensor, verify correct operation using the following checks:"),
  spacer(80),
  bulletPara("At 360 mm: Black wire output (OUT1) = ON → Jam Sensor active.", "bullets5"),
  bulletPara("At 530 mm: White wire output (OUT2) = ON → Bucket Door Sensor active.", "bullets5"),
  bulletPara("Out of both detection ranges: Both outputs OFF.", "bullets5"),
  spacer(80),
  bodyPara("You can verify the sensor outputs in real time using the I/O Control Panel on the HMI (see Sections 4.4 through 4.8)."),
  spacer(160),

  subHeading("5.6 Installer Notes", "s5-6"),
  bulletPara("Clean the sensor lens before powering up. Dust or debris on the lens can cause false readings or missed detections.", "bullets5"),
  bulletPara("If readings are unstable or inconsistent, check the sensor's physical alignment and ensure there is no excessive ambient light interference in the detection zone.", "bullets5"),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),

  // ── Pneumatic Sensor Assembly ──────────────────────────────────────────────
  subHeading("5.7 Pneumatic Sensor Assembly", "s5-7"),
  bodyPara("The machine includes a stacked pneumatic sensor assembly mounted on a DIN rail bracket. The assembly contains four SMC digital sensors arranged from top to bottom in the following order:"),
  spacer(80),
  bulletPara("Position 1 (Top) – Pressure Sensor for WB1  |  SMC ISE20A-V", "bullets5"),
  bulletPara("Position 2 – Vacuum Sensor for WB1  |  SMC ZSE20B-T", "bullets5"),
  bulletPara("Position 3 – Pressure Sensor for WB2  |  SMC ISE20A-V", "bullets5"),
  bulletPara("Position 4 (Bottom) – Vacuum Sensor for WB2  |  SMC ZSE20B-T", "bullets5"),
  spacer(120),
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
  spacer(80),
  bodyPara("Each sensor is wired independently to the machine's control panel. The assembly is located above the control panel. Refer to the electrical schematics for terminal assignments. Both sensor types share the same 5-wire color-coded wiring standard and 3-screen display interface, but monitor different pressure ranges."),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),

  // ── ZSE20B-T Vacuum Sensor ─────────────────────────────────────────────────
  subHeading("5.8 SMC ZSE20B-T – Digital Vacuum Sensor (WB1 & WB2)", "s5-8"),
  bodyPara("The SMC ZSE20B-T is a digital vacuum pressure switch used to monitor the vacuum level at WB1 and WB2. It features a 3-color, 3-screen LCD display and a PNP switch output for the configured setpoint."),
  spacer(120),

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
  spacer(160),

  subHeading("5.8.2 Setting SP1 – ZSE20B-T"),
  bodyPara("Use the front panel buttons to configure SP1 to the required vacuum level for your application:"),
  spacer(80),
  bulletPara("Step 1 – Press and hold the SET button for 3 seconds to enter setting mode. The display will show 'SP1'.", "bullets5"),
  bulletPara("Step 2 – Use the UP (▲) and DOWN (▼) buttons to adjust the SP1 value to the desired vacuum level. Press SET to confirm.", "bullets5"),
  bulletPara("Step 3 – The sensor returns to run mode and displays the live vacuum reading.", "bullets5"),
  spacer(80),
  infoBox("ℹ NOTE:", "Display color changes based on output state: green = output ON (vacuum reached SP1), red = output OFF (no vacuum detected). The display unit can be changed by pressing UP and DOWN simultaneously while in run mode.", LIGHT_BLUE),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),

  // ── ISE20A-V Pressure Sensor ───────────────────────────────────────────────
  subHeading("5.9 SMC ISE20A-V – Digital Pressure Sensor (WB1 & WB2)", "s5-9"),
  bodyPara("The SMC ISE20A-V is a digital pressure switch used to monitor pneumatic pressure at both bagging stations (WB1 and WB2). Three units are installed on the assembly — one for WB1 and two for WB2. Each features a 3-color, 3-screen LCD display and a PNP switch output for the configured setpoint."),
  spacer(120),

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
  spacer(160),

  subHeading("5.9.2 Changing the Display Unit to PSI – ISE20A-V"),
  bodyPara("Before setting the setpoint, confirm the sensor is displaying in PSI. If the display shows MPa, kPa, or another unit, follow these steps to switch to PSI:"),
  spacer(80),
  bulletPara("Step 1 – While in run mode, press and hold both the UP (▲) and DOWN (▼) buttons simultaneously for 3 seconds until the unit selection screen appears.", "bullets5"),
  bulletPara("Step 2 – Use UP/DOWN to scroll through the available units: MPa → kPa → kgf/cm² → bar → psi.", "bullets5"),
  bulletPara("Step 3 – When 'psi' is shown, press SET to confirm. The display will return to run mode showing the live reading in PSI.", "bullets5"),
  spacer(160),

  subHeading("5.9.3 Setting SP1 – ISE20A-V"),
  bodyPara("SP1 is set to 75 PSI. Although the main air supply regulator is set to 90 PSI, the sensor threshold is intentionally set lower to account for the pressure drop that occurs when the vacuum generator activates — preventing false low-pressure faults during normal operation. Do not raise SP1 above 75 PSI."),
  spacer(80),
  bulletPara("Step 1 – Press and hold the SET button for 3 seconds to enter setting mode. The display will show 'SP1'.", "bullets5"),
  bulletPara("Step 2 – Use the UP (▲) and DOWN (▼) buttons to set SP1 to 75 PSI. Press SET to confirm.", "bullets5"),
  bulletPara("Step 3 – The sensor returns to run mode and displays the live pressure reading in PSI.", "bullets5"),
  spacer(80),
  infoBox("ℹ NOTE:", "Display color changes based on output state: green = pressure at or above SP1 (normal), red = pressure below SP1 (fault condition). Response time can be adjusted in the sensor's function settings to reduce chattering from brief pressure fluctuations during valve switching.", LIGHT_BLUE),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 6. Fault Codes ───────────────────────────────────────────────────────────
const faultSection = [
  sectionHeading("6. Common Fault Codes & Troubleshooting", "s6"),
  spacer(100),
  bodyPara("If the machine displays a fault code or the ALARMS indicator illuminates in the Navigation Bar, refer to the table below. Tap ALARMS on the HMI to view active alarm details. For faults not listed here, contact the manufacturer's technical support."),
  spacer(120),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1440, 3000, 2520, 2400],
    rows: [
      new TableRow({
        children: ["Fault Code", "Description", "Possible Cause", "Corrective Action"].map((h, i) =>
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
            shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
            width: { size: [1440, 3000, 2520, 2400][i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText(h, 22, WHITE)] })],
          })
        ),
      }),
      ...["E01","E02","E03","E04","E05"].map((code, i) =>
        new TableRow({
          children: [code, "[Fault description]", "[Probable cause]", "[Corrective action]"].map((val, j) =>
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: i % 2 === 0 ? LIGHT_GRAY : WHITE, type: ShadingType.CLEAR },
              width: { size: [1440, 3000, 2520, 2400][j], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 80 },
              children: [new Paragraph({ children: [j === 0 ? boldText(code, 22, MID_BLUE) : placeholderRun(val, 22)] })],
            })
          ),
        })
      ),
    ],
  }),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 7. Contact & Support ─────────────────────────────────────────────────────
const contactSection = [
  sectionHeading("7. Contact & Technical Support"),
  spacer(100),
  bodyPara("For technical assistance, spare parts, or service, please contact your authorized service representative or the manufacturer directly:"),
  spacer(120),
  kvTable([
    ["Manufacturer:",       "[Manufacturer Name]"],
    ["Address:",            "[Street, City, State/Province, Country, ZIP]"],
    ["Phone:",              "[+X (XXX) XXX-XXXX]"],
    ["Email:",              "[support@manufacturer.com]"],
    ["Website:",            "[www.manufacturer.com]"],
    ["Service Hours:",      "[e.g., Mon–Fri, 8:00 AM – 5:00 PM (local time)]"],
    ["Parts Hotline:",      "[+X (XXX) XXX-XXXX]"],
    ["Local Distributor:",  "[Distributor name and contact]"],
  ]),
  spacer(200),
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
        ...contactSection,
      ],
    },
  ],
});

const outputPath = "/sessions/keen-wonderful-bohr/mnt/User-Manual/FSDWB4-Manual/Machine_User_Manual.docx";
const stream = Packer.toStream(doc);
const out = fs.createWriteStream(outputPath);
stream.pipe(out);
out.on("finish", () => console.log("✅ Document generated successfully."));
out.on("error", (err) => { console.error("❌ Error:", err); process.exit(1); });
t generated successfully."));
out.on("error", (err) => { console.error("Error:", err); process.exit(1); });
});
