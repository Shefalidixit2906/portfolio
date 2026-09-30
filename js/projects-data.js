// =========================================================
// PROJECTS DATA — this is the ONLY file you edit to add,
// remove, or update a project. No HTML or CSS changes needed.
//
// WHERE SCREENSHOTS/VIDEOS GO:
//   Put each project's media in its own folder, using lowercase
//   names with hyphens instead of spaces (spaces in file/folder
//   names cause broken links on some hosts):
//     assets/project/<project-name>/<file>.png (or .mp4)
//   e.g. assets/project/delegation-task/screen-1.png
//   (Older projects below still use the flat assets/*.svg files
//   from before this folder convention — both work fine.)
//
// HOW TO ADD A NEW PROJECT:
//   1. Add your screenshot(s)/video as described above.
//   2. Copy one of the { ... } blocks below (from the opening
//      "{" to its closing "}," — everything in between).
//   3. Paste your copy anywhere in the list (don't forget the
//      comma between blocks — every block except the LAST one
//      in the list needs a trailing comma after its "}").
//   4. Edit these fields in your new block:
//        image        -> path to the ONE screenshot used as the
//                         card thumbnail (see above). Optional if
//                         `video` is set — the video becomes the
//                         thumbnail instead (used as its poster
//                         frame if both are set).
//        imageAlt     -> short description of that screenshot
//        images       -> OPTIONAL list of ALL screenshots for this
//                         project (thumbnail included). When set,
//                         the popup shows every image in this list
//                         as a strip.
//        video        -> OPTIONAL path to a short .mp4 demo clip.
//                         Shown as an autoplaying muted loop on the
//                         card, and as a full player (with controls)
//                         at the top of the popup. Keep clips small —
//                         compress with ffmpeg (scale + fps + crf) if
//                         the raw screen recording is large; the
//                         Cowork Artifact host caps each file at 15MB.
//        title        -> project title
//        description  -> 1-2 sentence summary + impact/result
//        tags         -> list of tools/technologies used
//        previewLink  -> "#" if nothing to link yet, or the URL of
//                         your deployed Google Apps Script web app
//                         (or any live demo). Shows a separate
//                         "Live preview ↗" button that opens the
//                         URL in a new tab.
//        sheetLink    -> "#" if nothing to link yet, or the URL of
//                         the Google Sheet behind this project
//                         (Share → "Anyone with the link" → Viewer,
//                         then paste that link here). Shows a
//                         separate "View Sheet ↗" button that opens
//                         the sheet in a new tab.
//        caseStudy    -> OPTIONAL. Fill this in and the popup shows
//                         a full challenge/approach/result writeup
//                         (plus the video/images above, if any).
//                         Leave the whole caseStudy: {...} block out
//                         (or set it to null) to skip the writeup —
//                         the popup still opens on video/images alone.
//          challenge    -> the problem/situation before this project
//          approach     -> what you did, step by step
//          result       -> the outcome/impact (numbers if you have them)
//          learnings    -> OPTIONAL one-line takeaway; delete the
//                          line entirely if you don't want one
//        stats        -> OPTIONAL list of { value, label } pairs shown
//                         as a small stat row at the top of the popup —
//                         only use numbers/facts you can actually back up
//                         (e.g. { value: "60%", label: "Manual effort cut" }).
//                         Leave out entirely if you don't have a real one yet.
//   5. Save the file. Reload the page — that's it.
//
// TO REMOVE A PROJECT: delete its whole { ... } block below
// (and fix up commas so no two blocks are separated by more
// than one comma, and the last block has none after it).
//
// The order here is the display order on the page.
// =========================================================

window.PORTFOLIO_PROJECTS = [
  {
    image: "Vendor Collection & Payment Management System/assets/image/screen-1.png",
    imageAlt: "Outstanding Dashboard showing total outstanding, dealers with dues, ageing breakdown, and outstanding trend",
    images: [
      "Vendor Collection & Payment Management System/assets/image/screen-1.png",
      "Vendor Collection & Payment Management System/assets/image/screen-2.png",
      "Vendor Collection & Payment Management System/assets/image/screen-3.png"
    ],
    video: "Vendor Collection & Payment Management System/assets/Vedio/demo.mp4",
    title: "Vendor Collection & Payment Management System",
    description: "Centralized system automating payment tracking, collection status, and balance reconciliation, with role-based dashboards and automated reminders — cut manual data processing by 60%.",
    tags: ["Google Apps Script", "Google Sheets", "JavaScript"],
    previewLink: "https://script.google.com/macros/s/AKfycbxJIqDrU304aSUBZ724g_k21O1aI4iargDoGi7KiW2nbeyVKAebcGY1Jx6Ic09boxJAUw/exec",
    sheetLink: "https://docs.google.com/spreadsheets/d/1FPhlVlaDC67L1bNrNgf1Izxa8NHxLG-IAtVpbQWbvV8/edit?gid=0#gid=0",
    stats: [
      { value: "60%", label: "Manual effort cut" },
      { value: "Real-time", label: "MIS reporting" }
    ],
    caseStudy: {
      challenge: "Vendor payments and collections were tracked across scattered spreadsheets with manual follow-ups, making it hard to see real-time balances and leading to missed follow-ups and reconciliation errors.",
      approach: "Gathered requirements from finance and vendor-management stakeholders, mapped the existing manual process end-to-end, then designed a centralized system in Google Apps Script and Google Sheets with role-based dashboards, automated payment tracking, and scheduled email reminders for pending collections.",
      result: "Cut manual data processing effort by 60% through workflow automation. Real-time MIS reports replaced end-of-week manual reconciliation, and automated reminders removed the need to chase collections by hand.",
      learnings: "Role-based access was the key unlock — once each stakeholder only saw their relevant view, adoption went up and errors from editing the wrong sheet dropped sharply."
    }
  },
  {
    image: "assets/project-placeholder-2.svg",
    imageAlt: "Order Management System screenshot placeholder",
    title: "Order Management System (OMS)",
    description: "End-to-end order-to-production pipeline on Google Apps Script — auto-fills dealer and design data on entry, calculates production quantities from category-specific fabric ratios, flags overbooked designs automatically, and rolls everything into a live dashboard by dealer, executive, category, and supplier.",
    tags: ["Google Apps Script", "HTML/CSS", "JavaScript"],
    previewLink: "#",
    sheetLink: "#",
    caseStudy: {
      challenge: "Order entries had no live link to dealer/design master data or production capacity, so incomplete dealer info, mismatched design codes, and overbooked designs weren't caught until they'd already reached the production floor.",
      approach: "Built an Apps Script system that auto-fills dealer and design details the moment an order row is entered, calculates required pieces per design from category-specific fabric meter/ratio rules, and syncs every order into a production master sheet in real time. Added an automatic overbooking check that emails an alert the moment a design's booked quantity exceeds available capacity, plus a live web-app dashboard breaking orders down by dealer, executive, product category, and supplier.",
      result: "Removed manual piece-count calculation and cross-checking entirely, and overbooked designs are now caught and flagged the moment they happen instead of surfacing as a production shortfall.",
      learnings: "Putting the overbooking check into the data-sync step itself — rather than a separate manual review — is what actually closed the gap between order entry and the production floor."
    }
    // No real screenshot/video yet — the recording in this project's local
    // folder turned out to be a duplicate of the Task Delegation demo, so
    // it wasn't used here. Add a real one whenever it's ready.
  },
  {
    image: "assets/project/delegation-task/screen-1.png",
    imageAlt: "Delegation System dashboard — Follow Up Tasks tab with search and status table",
    images: [
      "assets/project/delegation-task/screen-1.png",
      "assets/project/delegation-task/screen-2.png",
      "assets/project/delegation-task/screen-3.png"
    ],
    video: "assets/project/delegation-task/demo.mp4",
    title: "Task Delegation & Follow-Up System",
    description: "Web app for assigning tasks to team members with a planned date, then tracking follow-ups and pending items from one dashboard instead of chasing status over chat and calls.",
    tags: ["Google Apps Script", "Google Sheets", "JavaScript"],
    previewLink: "https://script.google.com/macros/s/AKfycbwS4UxQuPXYJvL3kUyxKs4yIvbBxu4FApQ0NIWPvkVhZYaB6hZIfLGdphzpk-w2EuRvTw/exec",
    sheetLink: "https://docs.google.com/spreadsheets/d/1vxMnHMDULblhb5wWoCbjriDUWjb5cV0_12O0pWCrYV8/edit?gid=0#gid=0",
    caseStudy: {
      challenge: "Task delegation across the team happened over chat and calls, so there was no shared record of what was assigned, whether the planned date had been met, or what still needed following up.",
      approach: "Built a Google Apps Script web app with three views — Add New Task (assign a task to a \"doer\" with a planned date and description), Follow Up (tasks due for review, searchable by name or task), and All Pending Tasks (a full status board with a WhatsApp-contacted flag) — plus a revise flow for pushing a task's planned date when it slips.",
      result: "Task status, ownership, and follow-up state now live in one dashboard instead of scattered across chat threads, so nothing gets forgotten between assignment and completion.",
      learnings: "Separating \"Follow Up\" (things due today) from \"All Pending Tasks\" (everything open) turned out to matter more than expected — it's the difference between a daily action list and a full audit view."
    }
  },
  {
    image: "assets/project-placeholder-5.svg",
    imageAlt: "Ultimate Checklist v7 screenshot placeholder",
    title: "Ultimate Checklist v7 — Recurring Task Automation",
    description: "Google Apps Script system that generates recurring staff tasks from a working-day calendar, assigns backup \"buddies\" for coverage, sends reminders, and archives completed items — replacing a manually re-typed weekly checklist.",
    tags: ["Google Apps Script", "Google Sheets", "JavaScript"],
    previewLink: "#",
    sheetLink: "#",
    caseStudy: {
      challenge: "Recurring staff tasks (daily/weekly checklist items) were maintained by manually re-typing them into a sheet each cycle, with no automatic handling of Sundays, holidays, or a backup person when the assigned doer was unavailable.",
      approach: "Built a versioned (v7) Apps Script system that reads a master task list and a working-day calendar, then auto-generates each cycle's due tasks while skipping non-working days. Added an \"Assign Buddy\" dialog so a backup owner can cover a doer's tasks, automatic email reminders and doer notifications, and an archive step that clears completed items from the active view.",
      result: "Recurring tasks now generate themselves against the actual working calendar instead of being retyped, and coverage gaps (someone on leave) are handled by explicit buddy assignment rather than tasks silently going undone.",
      learnings: "Building the working-day calendar as its own sheet — rather than hardcoding Sunday-only skipping — is what let it flex for holidays without touching the code."
    }
  },
  {
    image: "assets/project-placeholder-6.svg",
    imageAlt: "FMS Order-to-Delivery step tracker screenshot placeholder",
    title: "FMS — Order-to-Delivery Step Tracker",
    description: "No-code Google Sheets add-on (\"BMP Formulas\") that lets you define an order-to-delivery pipeline step by step through wizard dialogs, auto-generating the turn-around-time formulas, conditional formatting, and linked intake forms each step needs — without writing a single spreadsheet formula by hand.",
    tags: ["Google Apps Script", "Google Sheets", "JavaScript"],
    previewLink: "#",
    sheetLink: "#",
    caseStudy: {
      challenge: "Tracking an order through each stage from booking to delivery meant hand-writing turn-around-time (TAT) formulas, lead-time offsets, and conditional formatting for every new step — slow to build and easy to break when a step's logic changed.",
      approach: "Built a wizard-driven Apps Script add-on (\"BMP Formulas\") with dialogs to set up the pipeline, add the first step (which also creates its entry form), and add each subsequent step. Each wizard auto-generates the right formula type for that step — TAT, a fixed lead-time offset, a specific-time rule, or a conditional \"show if Yes/No\" — plus the matching conditional formatting and, where needed, a linked Google Form for data entry.",
      result: "Adding or changing a pipeline step is now a wizard dialog instead of a formula-writing exercise, and every step gets consistent TAT tracking and formatting automatically.",
      learnings: "Making \"Add First Step\" also generate the entry form solved the chicken-and-egg problem of needing a form before there was anything to track — the first step effectively bootstraps the rest of the pipeline."
    }
  },
  {
    video: "assets/project/sales-dashboard/demo.mp4",
    imageAlt: "Dealer Sales Performance Dashboard",
    title: "Dealer Sales Performance Dashboard",
    description: "Google Apps Script web app pulling live sales data from a spreadsheet to show KPI cards, a category-by-executive heatmap, and a searchable dealer table — covering roughly 161 dealers across 7 product categories managed by 5 executives.",
    tags: ["Google Apps Script", "Google Sheets", "JavaScript"],
    previewLink: "#",
    sheetLink: "#",
    stats: [
      { value: "161", label: "Dealers tracked" },
      { value: "7", label: "Categories" },
      { value: "5", label: "Executives" }
    ],
    caseStudy: {
      challenge: "Sales performance across ~161 dealers, 7 product categories, and 5 executives lived in separate monthly sales sheets with no consolidated view, so spotting which categories or executives were under-performing meant manually cross-referencing multiple sheets.",
      approach: "Built an Apps Script web app that reads a dealer/executive master sheet plus every monthly sales sheet, and renders it as KPI cards, a category-by-executive heatmap, and a searchable, filterable dealer table, using `HtmlService` with client-side `google.script.run` calls for live data.",
      result: "Sales performance across the full dealer network is now visible in one dashboard instead of requiring a manual roll-up across monthly sheets.",
      learnings: "Building the dealer-to-executive mapping as its own lookup step, separate from the sales rows, made it possible to re-attribute a dealer to a different executive without touching historical sales data."
    }
  },
  {
    image: "assets/project-placeholder-3.svg",
    imageAlt: "Sales Executive Performance Tracker screenshot placeholder",
    title: "Sales Executive Management & Performance Tracker",
    description: "Automated activity tracking, lead management, and incentive reporting with role-based access, cutting manual effort by 70%.",
    tags: ["Google Apps Script", "Google Sheets"],
    previewLink: "#",
    sheetLink: "#",
    caseStudy: {
      challenge: "Sales activity, lead follow-ups, and incentive calculations were tracked manually per executive, which made performance reviews slow and incentive payouts prone to error and dispute.",
      approach: "Built a centralized tracker with role-based access: each sales executive logged their own activity and leads, while managers got aggregated KPI dashboards and automated incentive summaries generated straight from that data.",
      result: "Reduced manual reporting effort by 70% and gave managers same-day visibility into targets vs. actuals instead of waiting for end-of-month reconciliation.",
      learnings: "Automating the incentive calculation removed a recurring point of dispute between sales reps and management — the numbers were transparent and traceable to the source data."
    }
  },
  {
    image: "assets/project-placeholder-4.svg",
    imageAlt: "Workassist requirements and UX delivery screenshot placeholder",
    title: "End-to-End Requirement & UX Delivery — Workassist",
    description: "Gathered functional and business requirements into a full SRS, guided Figma mockups via knowledge-transfer sessions, and drove UAT to acceptance.",
    tags: ["BRD", "Figma", "SRS", "UAT"],
    previewLink: "#",
    sheetLink: "#",
    caseStudy: {
      challenge: "Stakeholder requirements for a new feature needed to be translated into something both engineering and design could act on directly, without repeated rounds of back-and-forth clarification.",
      approach: "Ran requirement-gathering sessions with stakeholders and documented everything into a comprehensive SRS, then led knowledge-transfer sessions with the UI/UX team to guide Figma mockups directly from that document.",
      result: "Drove UAT participation through to full sign-off, with the shipped product matching every acceptance criterion and minimal rework requested after handoff.",
      learnings: "Walking the UI/UX team through the SRS in a working session — rather than just handing over a document — cut mockup revision cycles down noticeably."
    }
  }
];
