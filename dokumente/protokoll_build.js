const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, ShadingType: ST,
  Header, Footer, PageNumber, LevelFormat, convertMillimetersToTwip, TabStopType,
} = require("docx");
const fs = require("fs");

const ACCENT = "1F3864";   // dunkles Blau
const ACCENT_L = "8496B0";
const GRAY = "595959";
const LIGHT = "EDF0F5";
const RULE = "C7CEDB";

const CONTENT_W = 9355; // 16,5 cm in DXA

const NONE = { style: BorderStyle.NONE, size: 0, color: "auto" };

const p = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: 276 },
    alignment: opts.alignment,
    indent: opts.indent,
    children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size ?? 21, color: opts.color ?? "1A1A1A" })],
  });

// Absatz mit fettem Lead-in ("Label: Text")
const leadIn = (label, text) =>
  new Paragraph({
    spacing: { after: 140, line: 276 },
    children: [
      new TextRun({ text: label + ": ", bold: true, size: 21, color: ACCENT }),
      new TextRun({ text, size: 21, color: "1A1A1A" }),
    ],
  });

const bullet = (text, level = 0) =>
  new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 60, line: 276 },
    children: [new TextRun({ text, size: 21, color: "1A1A1A" })],
  });

const sectionHeading = (num, text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT_L, space: 6 } },
    children: [
      new TextRun({ text: num + "  ", bold: true, size: 26, color: ACCENT_L }),
      new TextRun({ text, bold: true, size: 26, color: ACCENT }),
    ],
  });

const cell = (children, { width, shading, bold, align, colSpan } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    columnSpan: colSpan,
    shading: shading ? { type: ShadingType.CLEAR, fill: shading, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: Array.isArray(children)
      ? children
      : [new Paragraph({
          alignment: align,
          spacing: { after: 0, line: 260 },
          children: [new TextRun({ text: children, bold, size: 20, color: bold ? ACCENT : "1A1A1A" })],
        })],
  });

const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
const gridBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  left: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: RULE },
};

/* ---------- Kopfdaten-Tabelle ---------- */
const metaRow = (label, value) =>
  new TableRow({
    children: [
      cell(label, { width: 2600, bold: true }),
      cell(value, { width: 6755 }),
    ],
  });

const metaTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2600, 6755],
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    left: NONE, right: NONE,
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideVertical: NONE,
  },
  rows: [
    metaRow("Gremium", "Vorstand CPC – Taunus e.V."),
    metaRow("Datum", "Montag, 31. August 2026"),
    metaRow("Beginn", "18:00 Uhr"),
    metaRow("Teilnehmer", "Alex, Axel, Joe"),
  ],
});

/* ---------- Terminoptionen MV ---------- */
const termine = [
  ["Option A", "28. September 2026", "Montag"],
  ["Option B", "29. September 2026", "Dienstag"],
  ["Option C", "01. Oktober 2026", "Donnerstag"],
  ["Option D", "02. Oktober 2026", "Freitag"],
];
const W_T = [1800, 4155, 3400];
const termineTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: W_T,
  borders: gridBorders,
  rows: [
    new TableRow({
      tableHeader: true,
      children: ["Option", "Datum", "Wochentag"].map((t, i) =>
        cell(t, { width: W_T[i], shading: LIGHT, bold: true })),
    }),
    ...termine.map((r) => new TableRow({
      children: r.map((t, i) => cell(t, { width: W_T[i], bold: i === 0 })),
    })),
  ],
});

/* ---------- Maßnahmenübersicht ---------- */
const massnahmen = [
  ["1", "Rückforderung der unautorisierten Abflüsse in Höhe von 3.399,50 € gegenüber Paul Stuhr", "Axel"],
  ["2", "Entzug der Kontovollmacht für Paul Stuhr (Antrag gestellt) – Nachverfolgung bei der Bank", "Vorstand"],
  ["3", "Bereitstellung und Einführung der App für den Beschaffungsprozess (Vier-Augen-Prinzip)", "Alex"],
  ["4", "Anfrage des Clubraums in Bad Soden bei David Hauser", "Axel"],
  ["5", "Terminabstimmung zur Mitgliederversammlung in der Stammgruppe", "Vorstand"],
  ["6", "Erarbeitung eines Vorschlags für die differenzierte Beitragsordnung", "Axel"],
  ["7", "Versand der Einladung zur Mitgliederversammlung unter Einhaltung der zweiwöchigen Frist", "Vorstand"],
];
const W_M = [700, 6855, 1800];
const massnahmenTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: W_M,
  borders: gridBorders,
  rows: [
    new TableRow({
      tableHeader: true,
      children: ["Nr.", "Maßnahme", "Verantwortlich"].map((t, i) =>
        cell(t, { width: W_M[i], shading: LIGHT, bold: true, align: i === 0 ? AlignmentType.CENTER : undefined })),
    }),
    ...massnahmen.map((r) => new TableRow({
      children: r.map((t, i) => cell(t, { width: W_M[i], align: i === 0 ? AlignmentType.CENTER : undefined })),
    })),
  ],
});

/* ---------- Unterschriftenblock ---------- */
const sigLine = (label) => [
  new Paragraph({
    spacing: { before: 600, after: 40 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "8C8C8C", space: 2 } },
    children: [new TextRun({ text: "", size: 21 })],
  }),
  new Paragraph({
    spacing: { after: 0 },
    children: [new TextRun({ text: label, size: 18, color: GRAY })],
  }),
];

const W_S = [4400, 555, 4400];
const signatureTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: W_S,
  borders: noBorders,
  rows: [
    new TableRow({
      children: [
        cell(sigLine("Sitzungsleitung"), { width: W_S[0] }),
        cell([new Paragraph({ children: [new TextRun({ text: "", size: 21 })] })], { width: W_S[1] }),
        cell(sigLine("Protokollführung"), { width: W_S[2] }),
      ],
    }),
  ],
});

/* ---------- Dokument ---------- */
const doc = new Document({
  creator: "CPC – Taunus e.V.",
  title: "Protokoll der Vorstandssitzung vom 31. August 2026",
  description: "Protokoll der Vorstandssitzung des CPC – Taunus e.V. vom 31. August 2026",
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 21, color: "1A1A1A" } },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Calibri", size: 26, bold: true, color: ACCENT },
        paragraph: { spacing: { before: 360, after: 160 } },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "–",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 240 } }, run: { color: ACCENT_L, bold: true } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 1418, right: 1134, bottom: 1134, left: 1418, header: 720, footer: 567 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              spacing: { after: 0 },
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 4 } },
              tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
              children: [
                new TextRun({ text: "CPC – Taunus e.V.", bold: true, size: 18, color: ACCENT, characterSpacing: 20 }),
                new TextRun({ text: "\tVorstandssitzung 31.08.2026", size: 18, color: GRAY }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              spacing: { before: 0 },
              border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
              tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
              children: [
                new TextRun({ text: "Vertraulich – nur für den internen Gebrauch", size: 16, color: GRAY }),
                new TextRun({ text: "\tSeite ", size: 16, color: GRAY }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY }),
                new TextRun({ text: " von ", size: 16, color: GRAY }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: GRAY }),
              ],
            }),
          ],
        }),
      },
      children: [
        // Titelblock
        new Paragraph({
          spacing: { before: 240, after: 40 },
          children: [new TextRun({ text: "PROTOKOLL", bold: true, size: 20, color: ACCENT_L, characterSpacing: 60 })],
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "Vorstandssitzung CPC – Taunus e.V.", bold: true, size: 40, color: ACCENT })],
        }),
        new Paragraph({
          spacing: { after: 320 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 8 } },
          children: [new TextRun({ text: "31. August 2026, 18:00 Uhr", size: 22, color: GRAY })],
        }),

        metaTable,

        // 1
        sectionHeading("1.", "Finanzwesen & Buchhaltung"),
        leadIn("Statusbericht", "Alex präsentierte die aktuelle Finanzübersicht. Die Buchhaltung wurde erfolgreich in eine neue Struktur (Excel/Pivot) überführt, um volle Transparenz über Einnahmen und Ausgaben zu gewährleisten."),
        leadIn("Unregelmäßigkeiten", "Es wurden unautorisierte Abflüsse in Höhe von 3.399,50 € identifiziert. Axel wird Paul Stuhr auffordern, diese Beträge umgehend zurückzuüberweisen, da keine entsprechenden Vorstandsbeschlüsse oder Belege vorliegen."),
        leadIn("Bankzugang", "Der Antrag auf Entzug der Kontovollmacht für Paul Stuhr wurde gestellt. Zukünftig wird das Konto primär durch den Vorstand (Axel, Alex, Joe und Stephan) geführt."),

        // 2
        sectionHeading("2.", "Einführung des Vier-Augen-Prinzips (App-Lösung)"),
        leadIn("Implementierung", "Zur Vermeidung künftiger Interessenskonflikte wird eine von Alex entwickelte App für den Beschaffungsprozess eingeführt."),
        leadIn("Workflow", "Alle Ausgaben (Bestellanforderungen, Rechnungen, Verträge) müssen zwingend ein digitales Genehmigungsverfahren durchlaufen."),
        leadIn("Zugang", "Der Vorstand erhält administrative Freigaberechte; die App steht kurz vor dem Release im App Store (Review-Phase ca. 48 Stunden), ist aber bereits als Web-App nutzbar."),

        // 3
        sectionHeading("3.", "Vorbereitung der Mitgliederversammlung (MV)"),
        leadIn("Terminplanung", "Es wurden vier potenzielle Termine zur Abstimmung in der Stammgruppe definiert:"),
        termineTable,
        new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "", size: 12 })] }),
        leadIn("Ort", "Clubraum in Bad Soden (Anfrage bei David Hauser durch Axel)."),
        new Paragraph({
          spacing: { after: 100 },
          children: [new TextRun({ text: "Agenda-Punkte (unter anderem):", bold: true, size: 21, color: ACCENT })],
        }),
        bullet("Beschlussfassung über die neue Geschäftsordnung der Mitgliederversammlung."),
        bullet("Satzungsänderung zur Handlungsfähigkeit des Aufsichtsrats."),
        bullet("Festlegung eines Budgets für Vereinswaffen und Training (Vorschlag: 5.000 € p. a.)."),
        bullet("Nachwahl eines Aufsichtsratsmitglieds (Nachfolge Maxi)."),

        // 4
        sectionHeading("4.", "Beitragsordnung & Mitgliederstruktur"),
        leadIn("Strategie", "Der Verein soll seinen exklusiven Charakter behalten."),
        leadIn("Anpassungen", "Axel erarbeitet einen Vorschlag für eine differenzierte Beitragsordnung (u. a. Regelungen für Familienmitglieder, Werbeprämien und Beitragsbefreiungen für Trainer und Funktionsträger). Diese wird vor der Mitgliederversammlung im Vorstand final abgestimmt."),

        // 5 Maßnahmen
        sectionHeading("5.", "Maßnahmenübersicht"),
        massnahmenTable,

        // Anmerkung
        new Paragraph({
          spacing: { before: 360, after: 0 },
          shading: { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" },
          border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT_L, space: 8 } },
          indent: { left: 120, right: 120 },
          children: [
            new TextRun({ text: "Anmerkung: ", bold: true, size: 20, color: ACCENT }),
            new TextRun({
              text: "Dieses Protokoll dient als Grundlage für die offizielle Einladung zur Mitgliederversammlung, die unter Einhaltung der zweiwöchigen Frist versendet wird.",
              size: 20, color: "1A1A1A", italics: true,
            }),
          ],
        }),

        signatureTable,
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "Protokoll_Vorstandssitzung_CPC-Taunus_2026-08-31.docx", buf);
  console.log("geschrieben");
});
