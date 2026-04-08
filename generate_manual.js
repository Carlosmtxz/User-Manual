const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
  LevelFormat, ImageRun
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
const IMG_DIR = "/sessions/wonderful-nice-faraday/mnt/FSDWB4-Manual/FSDWB4_HMI_ScreenShots/";
const LOGO_PATH = "/sessions/wonderful-nice-faraday/fox_logo.png";
// Fit ~820x500 HMI images within 6.5" content width at 96 DPI
const IMG_W = 624;
const IMG_H = 380;
// Logo: 346x146px, display at 3 inches wide on cover (288px @ 96dpi)
const LOGO_W = 288;
const LOGO_H = 122;

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

const sectionHeading = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: WHITE })],
    shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
    spacing: { before: 400, after: 200 },
    indent: { left: 180 },
  });

const subHeading = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: DARK_BLUE })],
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 2 } },
  });

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

// ─── Cover page ───────────────────────────────────────────────────────────────
const coverPage = [
  spacer(1600),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(LOGO_PATH),
      transformation: { width: LOGO_W, height: LOGO_H },
      altText: { title: "Fox Solutions Logo", description: "Fox Solutions company logo", name: "FoxSolutionsLogo" },
    })],
    spacing: { after: 240 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [normalText("Dual Wicketed Bagger User Manual", 36, MID_BLUE)],
    spacing: { after: 80 },
  }),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
        shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 200, bottom: 200, left: 300, right: 300 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [boldText("FSDWB4 – Dual Wicketed Bagger", 32, WHITE)] })],
      })],
    })],
  }),
  spacer(400),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [normalText("Version [X.X]  |  Document No. [DOC-XXXX]  |  Issue Date: [YYYY-MM-DD]", 22, GRAY)],
    spacing: { after: 800 },
  }),
  kvTable([
    ["Machine Name:",   "Dual Wicketed Bagger"],
    ["Model Number:",   "FSDWB4"],
    ["Manufacturer:",   "Fox Solutions"],
    ["Document Rev.:",  "[Enter revision number]"],
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
      ...["1.0", "1.1", "1.2"].map((rev, i) =>
        new TableRow({
          children: [rev, "[YYYY-MM-DD]", "[Describe change]", "[Approver]"].map((val, j) =>
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
const tocSection = [
  sectionHeading("Table of Contents"),
  spacer(100),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 1. Introduction ──────────────────────────────────────────────────────────
const introSection = [
  sectionHeading("1. Introduction"),
  spacer(100),
  subHeading("1.1 Purpose of This Manual"),
  bodyPara("This manual provides comprehensive operating instructions for the [Machine Name], Model [Model Number]. It is intended for use by qualified operators, technicians, and supervisors responsible for the day-to-day operation of the machine. Read all sections thoroughly before operating the equipment.", false),
  spacer(100),
  subHeading("1.2 Machine Overview"),
  bodyPara("The FSDWB4 Dual Wicketed Bagger is an automatic bagging machine designed for high-speed packaging of portioned products. The machine operates two bagging stations (WB1 and WB2), each connected to a single weigher that has multiple outlets. Each station sends a request signal to its corresponding weigher outlet when it is ready for product. Once the weigher delivers the correct portion, the station bags the product, seals the wicketed bag, and drops the finished bag onto an outfeed conveyor belt for downstream handling. The FSDWB4 is rated for up to 24 bags per minute at 1 lb portions."),
  spacer(100),
  subHeading("1.3 Intended Users"),
  bodyPara("This manual is intended for:"),
  bulletPara("Production operators responsible for running the FSDWB4 bagging machine during shifts", "bullets1"),
  bulletPara("Maintenance technicians performing routine upkeep or troubleshooting on the machine", "bullets1"),
  bulletPara("Production supervisors overseeing bagging line operations", "bullets1"),
  bulletPara("Quality assurance personnel monitoring bag weight and seal integrity", "bullets1"),
  spacer(100),
  subHeading("1.4 Technical Specifications"),
  spacer(80),
  kvTable([
    ["Machine Model:",        "FSDWB4"],
    ["Machine Type:",         "Dual Station Automatic Wicketed Bagger"],
    ["Number of Stations:",   "2 (WB1 and WB2)"],
    ["Rated Speed:",          "Up to 24 bags per minute (combined)"],
    ["Portion Size:",         "1 lb (standard)"],
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
  sectionHeading("2. Safety Overview"),
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
  sectionHeading("3. HMI Screen Reference"),
  spacer(100),
  bodyPara("This section provides a visual reference for all screens available on the Human-Machine Interface (HMI) touchscreen. Each screen is shown with a description of its purpose and key elements."),
  spacer(160),

  // 3.1 Main Screen
  subHeading("3.1 Main Screen"),
  bodyPara("The Main Screen is the primary operating interface displayed during normal production. It provides real-time status and control for both bagging stations (WB1 and WB2). From this screen, operators can monitor the bagging cycle, enable or disable individual station subsystems, and start or stop production."),
  spacer(120),
  hmiImage("Main_Page.png", "Main Screen"),
  imgCaption("Figure 3.1 – Main Screen (Normal Operation)"),
  spacer(80),
  bodyPara("Key elements of the Main Screen:"),
  bulletPara("Navigation Bar (top): HOME, SETUP, PARAMETERS, RECIPES, ALARMS — provides access to all major screens.", "bullets3"),
  bulletPara("START / ESTOP Buttons (top right): Green START initiates production; red ESTOP immediately halts all motion.", "bullets3"),
  bulletPara("Active Recipe Display: Shows the currently selected recipe (e.g., 1-2LB) and machine run status.", "bullets3"),
  bulletPara("WB1 / WB2 Panels: Each bagging station has independent controls including WEIGHER DUMP OFF and TEST DUMP buttons.", "bullets3"),
  bulletPara("KWIKLOK: Enables or disables the Kwiklok bag closure device for the respective bagging station.", "bullets3"),
  bulletPara("ENABLE WB: Activates or deactivates the bagging station from participating in the production cycle.", "bullets3"),
  bulletPara("AIR SUPPLY: Monitors or controls the pneumatic air supply status for each station.", "bullets3"),
  spacer(200),

  // 3.2 Recipes Screen
  subHeading("3.2 Recipes Screen"),
  bodyPara("The Recipes Screen is accessed from the RECIPES tab in the Navigation Bar. It allows operators to select the active weight recipe for the current production run. Each recipe corresponds to a product weight range and configures the machine timing and parameters accordingly."),
  spacer(120),
  hmiImage("Recipes_Page.png", "Recipes Screen"),
  imgCaption("Figure 3.2 – Recipes Screen"),
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
  spacer(200),

  // 3.3 Alarms Screen
  subHeading("3.3 Alarms Screen"),
  bodyPara("The Alarms Screen is accessed from the ALARMS tab in the Navigation Bar. It displays a log of all active and historical alarms, including the alarm number, descriptive message, and activation time. Operators should review this screen when the machine stops unexpectedly or when an alarm indicator is visible."),
  spacer(120),
  hmiImage("Alarms_Page.png", "Alarms Screen"),
  imgCaption("Figure 3.3 – Alarms Screen"),
  spacer(80),
  bodyPara("Alarm log columns:"),
  bulletPara("Alarm No – A unique identifier for each alarm condition.", "bullets3"),
  bulletPara("Message – A plain-text description of the fault or warning condition.", "bullets3"),
  bulletPara("Activated – The date and time the alarm was triggered.", "bullets3"),
  spacer(80),
  infoBox("⚠ WARNING:", "Do not clear alarms without first identifying and resolving the root cause. Repeatedly clearing an active alarm without fixing the fault may result in equipment damage or injury.", "FFF3CD"),
  spacer(200),

  // 3.4 Setup Menu – Page 1
  subHeading("3.4 Setup Menu – Page 1"),
  bodyPara("The Setup Menu is accessed by tapping the SETUP tab in the Navigation Bar. Page 1 contains diagnostic tools and system configuration options. Some functions require a technician login to access."),
  spacer(120),
  hmiImage("Setup_Page.png", "Setup Menu Page 1"),
  imgCaption("Figure 3.4 – Setup Menu, Page 1"),
  spacer(80),
  bodyPara("Available options on Setup Page 1:"),
  bulletPara("IO – View real-time input and output signal states; used for wiring verification and diagnostics (see Sections 3.8 and 3.9).", "bullets4"),
  bulletPara("TEST PARTS – Manually actuate individual pneumatic and mechanical components for testing (see Section 3.6).", "bullets4"),
  bulletPara("SYSTEM INFO – Display CPU program, firmware version, and production counters.", "bullets4"),
  bulletPara("CLEAN UP – Initiate a cleanup or purge cycle to clear product from the machine.", "bullets4"),
  bulletPara("OPTIONS – Configure sensor bypass toggles and feature enable/disable settings.", "bullets4"),
  bulletPara("LANGUAGE – Change the HMI display language.", "bullets4"),
  bulletPara("SIMULATION – Enable simulation mode for testing machine logic without live product or signals.", "bullets4"),
  spacer(200),

  // 3.5 Setup Menu – Page 2
  subHeading("3.5 Setup Menu – Page 2"),
  bodyPara("Setup Page 2 provides access to advanced diagnostic and drive configuration tools. These are primarily used by maintenance technicians during commissioning or troubleshooting."),
  spacer(120),
  hmiImage("Setup_Page2.png", "Setup Menu Page 2"),
  imgCaption("Figure 3.5 – Setup Menu, Page 2"),
  spacer(80),
  bodyPara("Available options on Setup Page 2:"),
  bulletPara("SYNC TEST – Tests the communication and synchronization between the bagging stations (WB1 and WB2) and the connected weigher. Use this after initial setup or when troubleshooting request/dump signal issues.", "bullets5"),
  bulletPara("DRIVES – Access individual drive parameter screens for WB1 and WB2 feed belt motors (see Sections 3.13 through 3.16).", "bullets5"),
  spacer(200),

  // 3.6 Test Parts Screen
  subHeading("3.6 Test Parts Screen"),
  bodyPara("The Test Parts Screen is accessed via Setup Menu → TEST PARTS. It allows technicians to manually activate individual pneumatic cylinders and mechanical components for each bagging station (WB1 and WB2) independently. This screen is used during maintenance, commissioning, and troubleshooting to verify component operation without running a full production cycle."),
  spacer(120),
  hmiImage("TestParts_Page1.png", "Test Parts Screen"),
  imgCaption("Figure 3.6 – Test Parts Screen"),
  spacer(80),
  bodyPara("Each output on this screen can be triggered using one of two modes:"),
  bulletPara("LATCH – A single press activates the output and it remains ON. Press again to turn it OFF. Use this mode when you need to hold a component in position for inspection.", "bullets4"),
  bulletPara("PULSE – The output only stays ON while you press and hold the button (minimum 1 second). As soon as you release, the output turns back OFF. Use this for a controlled, momentary actuation.", "bullets4"),
  spacer(80),
  infoBox("⚠ WARNING:", "Keep hands and body clear of all moving parts when using the Test Parts screen. Components will actuate immediately when triggered.", "FFF3CD"),
  spacer(200),

  // 3.7 Parameters Screen
  subHeading("3.7 Parameters Screen"),
  bodyPara("The Parameters Screen is accessed via the PARAMETERS tab in the Navigation Bar (technician login required). It displays and allows adjustment of timing values for both bagging stations (WB1 and WB2) that govern the product fall and feed cycle."),
  spacer(120),
  hmiImage("Parameters_Page.png", "Parameters Screen"),
  imgCaption("Figure 3.7 – Parameters Screen"),
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
  bodyPara("Each parameter is independently configurable for WB1 and WB2 to accommodate differences between the two stations."),
  spacer(200),

  // 3.8 I/O – Inputs Page 1
  subHeading("3.8 I/O Control Panel – Inputs (Page 1)"),
  bodyPara("The I/O Control Panel is accessed via Setup Menu → IO. It displays the real-time state of all machine inputs and outputs, allowing technicians to verify wiring, diagnose faults, and confirm signal integrity without requiring a PLC programmer. Inputs Page 1 covers the first set of digital input signals."),
  spacer(120),
  hmiImage("Inputs1_Page.png", "I/O Inputs Page 1"),
  imgCaption("Figure 3.8 – I/O Control Panel – Inputs, Page 1"),
  spacer(80),
  bodyPara("Each row in the panel shows the signal name alongside a color-coded indicator: green indicates the signal is active (ON), grey indicates inactive (OFF). Use this screen to quickly verify sensor states and confirm that E-stop, safety, and position switches are functioning correctly before starting production."),
  spacer(200),

  // 3.9 I/O – Inputs Page 2
  subHeading("3.9 I/O Control Panel – Inputs (Page 2)"),
  bodyPara("Inputs Page 2 continues the digital input signal display, covering additional sensors and switches not shown on Page 1."),
  spacer(120),
  hmiImage("Inputs2_Page.png", "I/O Inputs Page 2"),
  imgCaption("Figure 3.9 – I/O Control Panel – Inputs, Page 2"),
  spacer(80),
  bodyPara("Review both input pages when troubleshooting unexpected machine behavior or verifying that all safety interlocks are responding correctly."),
  spacer(200),

  // 3.10 I/O – Outputs Page 1
  subHeading("3.10 I/O Control Panel – Outputs (Page 1)"),
  bodyPara("The Outputs pages display the real-time state of all digital output signals, including solenoid valves, motors, and other actuators controlled by the PLC. Page 1 covers the first group of output signals in the I/O Control Panel."),
  spacer(120),
  hmiImage("Outputs1_Page.png", "I/O Outputs Page 1"),
  imgCaption("Figure 3.10 – I/O Control Panel – Outputs, Page 1"),
  spacer(80),
  bodyPara("Outputs are shown with their current state (ON/OFF). During troubleshooting, technicians can use this screen to confirm whether the PLC is commanding a device to activate, which helps isolate whether a fault is in the control logic or in the physical wiring and hardware."),
  spacer(200),

  // 3.11 I/O – Outputs Page 2 (Pneumatic Panel)
  subHeading("3.11 I/O Pneumatic Panel – Outputs (Page 2)"),
  bodyPara("The Pneumatic Panel Outputs display shows the state of the pneumatic solenoid valves and related output signals. This page is used to verify that pneumatic actuators are being correctly commanded during the bagging cycle."),
  spacer(120),
  hmiImage("Outputs2_Page.png", "I/O Pneumatic Outputs Page 2"),
  imgCaption("Figure 3.11 – I/O Pneumatic Panel – Outputs, Page 2"),
  spacer(200),

  // 3.12 I/O – Outputs Page 3 (Pneumatic Panel cont'd)
  subHeading("3.12 I/O Pneumatic Panel – Outputs (Page 3)"),
  bodyPara("Page 3 of the Pneumatic Panel Outputs continues the display of pneumatic output signals for both WB1 and WB2 stations."),
  spacer(120),
  hmiImage("Outputs3_Page.png", "I/O Pneumatic Outputs Page 3"),
  imgCaption("Figure 3.12 – I/O Pneumatic Panel – Outputs, Page 3"),
  spacer(80),
  infoBox("ℹ NOTE:", "Output forcing (manually commanding an output from the HMI) may be available in certain modes. Only qualified technicians should use this function. Forcing outputs with the machine running can cause unexpected motion.", LIGHT_BLUE),
  spacer(200),

  // 3.13 WB1 Flat Belt Drive Parameters
  subHeading("3.13 Drive Parameters – WB1 Flat Belt"),
  bodyPara("The WB1 Flat Belt Drive Parameters screen is accessed via Setup → DRIVES. It displays and allows adjustment of the speed and acceleration settings for the WB1 flat belt conveyor motor. A TEST function is available to run the drive briefly for verification."),
  spacer(120),
  hmiImage("WB1_FlatDrive_Parameters.png", "WB1 Flat Belt Drive Parameters"),
  imgCaption("Figure 3.13 – WB1 Flat Belt Drive Parameters"),
  spacer(80),
  bulletPara("SPEED (Hz) – Operating frequency of the flat belt drive motor. Increase to run the belt faster.", "bullets8"),
  bulletPara("ACCEL (ms) – Ramp-up time from 0 Hz to the target speed. Lower values produce faster acceleration.", "bullets8"),
  bulletPara("DRIVE 1 STATUS – Indicates the health of the drive (green = OK, fault indicator if issue detected).", "bullets8"),
  bulletPara("TEST Button – Runs the drive momentarily at the configured speed for commissioning or verification.", "bullets8"),
  spacer(200),

  // 3.14 WB1 V-Drive Parameters
  subHeading("3.14 Drive Parameters – WB1 V-Drive (Feed Belt)"),
  bodyPara("The WB1 V-Drive Parameters screen shows the configuration for the WB1 Feed Belt variable speed drives. The feed belt has two speed zones — an outer high-RPM zone and an inner low-RPM zone — each driven independently."),
  spacer(120),
  hmiImage("WB1_VDrive_Parameters.png", "WB1 V-Drive Parameters"),
  imgCaption("Figure 3.14 – WB1 V-Drive Parameters (Feed Belt Outer / Inner)"),
  spacer(80),
  bulletPara("FEED BELT OUTER – HIGH RPM: Controls the outer section of the WB1 feed belt at high speed. Drive 1 status shown.", "bullets8"),
  bulletPara("FEED BELT INNER – LOW RPM: Controls the inner section of the WB1 feed belt at low speed. Drive 2 status shown.", "bullets8"),
  bulletPara("SPEED (Hz) and ACCEL (ms): Configurable for each zone independently.", "bullets8"),
  bulletPara("TEST Button: Runs each belt zone independently for verification.", "bullets8"),
  spacer(200),

  // 3.15 WB2 Flat Belt Drive Parameters
  subHeading("3.15 Drive Parameters – WB2 Flat Belt"),
  bodyPara("The WB2 Flat Belt Drive Parameters screen displays speed and acceleration settings for the WB2 flat belt conveyor motor, mirroring the WB1 configuration on a separate drive channel."),
  spacer(120),
  hmiImage("WB2_FlatDrive_Parameters.png", "WB2 Flat Belt Drive Parameters"),
  imgCaption("Figure 3.15 – WB2 Flat Belt Drive Parameters"),
  spacer(80),
  bulletPara("SPEED (Hz) – Operating frequency of the WB2 flat belt drive.", "bullets9"),
  bulletPara("ACCEL (ms) – Motor ramp-up time for the WB2 flat belt.", "bullets9"),
  bulletPara("DRIVE 4 STATUS – Health indicator for Drive 4 (WB2 flat belt).", "bullets9"),
  bulletPara("TEST Button – Activates the WB2 flat belt briefly for testing.", "bullets9"),
  spacer(200),

  // 3.16 WB2 V-Drive Parameters
  subHeading("3.16 Drive Parameters – WB2 V-Drive (Feed Belt)"),
  bodyPara("The WB2 V-Drive Parameters screen shows configuration for the WB2 Feed Belt variable speed drives. Like WB1, the WB2 feed belt has inner and outer speed zones controlled by separate drives (Drive 3 and Drive 4)."),
  spacer(120),
  hmiImage("WB2_VDrive_Parameters.png", "WB2 V-Drive Parameters"),
  imgCaption("Figure 3.16 – WB2 V-Drive Parameters (Feed Belt Inner / Outer)"),
  spacer(80),
  bulletPara("FEED BELT INNER – LOW RPM: Controls the inner section of the WB2 feed belt at low speed. Drive 3 status shown.", "bullets9"),
  bulletPara("FEED BELT OUTER – HIGH RPM: Controls the outer section of the WB2 feed belt at high speed. Drive 4 status shown.", "bullets9"),
  bulletPara("SPEED (Hz) and ACCEL (ms): Independently configurable for each belt zone.", "bullets9"),
  bulletPara("TEST Button: Runs each zone separately for commissioning verification.", "bullets9"),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 4. Operating Instructions ────────────────────────────────────────────────
const operatingSection = [
  sectionHeading("4. Operating Instructions"),
  spacer(100),
  infoBox("ℹ BEFORE YOU BEGIN:", "Ensure you have completed the pre-operation checklist in Section 4.1 before starting the machine. Refer to Section 3 for descriptions of all HMI screens referenced below.", LIGHT_BLUE),
  spacer(200),

  subHeading("4.1 Pre-Operation Checklist"),
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

  subHeading("4.2 Starting the Machine"),
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
  stepCard(5, "Start Production", "Press the green START button (top right of Main Screen). Each station will begin its cycle: requesting product from the weigher, receiving the 1 lb portion, bagging and sealing it, then dropping the finished bag onto the outfeed conveyor. The machine is rated for up to 24 bags per minute. Monitor the first several cycles to confirm normal operation on both stations."),
  spacer(200),

  subHeading("4.3 During Normal Operation"),
  bodyPara("While the machine is running, monitor the Main Screen continuously for the following:"),
  bulletPara("WEIGHER DUMP OFF buttons — confirm each bagging station is cycling correctly.", "bullets11"),
  bulletPara("Dynamic Text product name — confirm the correct recipe remains active.", "bullets11"),
  bulletPara("Watch for any unexpected button state changes or alarm indicators on the Navigation Bar.", "bullets11"),
  bulletPara("Check physical bag output quality at regular intervals.", "bullets11"),
  spacer(120),
  infoBox("ℹ TIP:", "If output quality changes unexpectedly, navigate to PARAMETERS to verify timing values have not been altered.", LIGHT_BLUE),
  spacer(200),

  subHeading("4.4 Stopping the Machine"),
  bodyPara("At the end of each production run:"),
  numberedPara("Press the STOP / pause control to end the cycle after the current bag is completed.", "numbers1"),
  numberedPara("Allow all moving components to come to a full stop before approaching the machine.", "numbers1"),
  numberedPara("Record production totals from the SYSTEM INFO screen (Current WB1/WB2 and Total counts).", "numbers1"),
  numberedPara("Toggle ENABLE WB and AIR SUPPLY to OFF for both bays.", "numbers1"),
  numberedPara("Clean the machine and surrounding area.", "numbers1"),
  numberedPara("Turn the main power switch to the OFF position.", "numbers1"),
  numberedPara("Complete the operator production log.", "numbers1"),
  spacer(200),

  subHeading("4.5 Emergency Stop Procedure"),
  infoBox("🛑 EMERGENCY:", "In any emergency, press the red ESTOP button at the top right of the Main Screen immediately. This cuts power to all motion systems.", "FDDEDE"),
  spacer(160),
  bodyPara("After activating the ESTOP:"),
  numberedPara("Do NOT attempt to restart until the cause of the emergency has been identified and resolved.", "numbers2"),
  numberedPara("Alert your supervisor and/or the designated safety officer immediately.", "numbers2"),
  numberedPara("If there is an injury, call [Emergency Number / First Aid contact] immediately.", "numbers2"),
  numberedPara("Complete an incident report before the machine is returned to service.", "numbers2"),
  numberedPara("To reset: twist the ESTOP button clockwise to release it, then follow the normal startup procedure from Section 4.2.", "numbers2"),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ─── 5. Fault Codes ───────────────────────────────────────────────────────────
const faultSection = [
  sectionHeading("5. Common Fault Codes & Troubleshooting"),
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

// ─── 6. Contact & Support ─────────────────────────────────────────────────────
const contactSection = [
  sectionHeading("6. Contact & Technical Support"),
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
        ...hmiSection,
        ...operatingSection,
        ...faultSection,
        ...contactSection,
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/sessions/wonderful-nice-faraday/mnt/FSDWB4-Manual/Machine_User_Manual.docx", buffer);
  console.log("✅ Document generated successfully.");
}).catch(err => {
  console.error("❌ Error:", err);
  process.exit(1);
});
