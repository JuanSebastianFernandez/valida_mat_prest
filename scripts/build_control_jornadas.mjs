import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const rootDir = process.cwd();
const seedPath = path.join(rootDir, "data", "jornada_seed.json");
const outputDir = path.join(rootDir, "outputs", "019f8bfd-f2ff-7d72-9352-33571d645b75");
const previewDir = path.join(outputDir, "previews");
const outputPath = path.join(outputDir, "control_jornadas.xlsx");
const seed = JSON.parse(await fs.readFile(seedPath, "utf8"));

const reportByPerson = new Map();
for (const report of seed.reports) {
  if (!reportByPerson.has(report.sender)) {
    reportByPerson.set(report.sender, { Inicio: [], Cierre: [] });
  }
  reportByPerson.get(report.sender)[report.type].push(report);
}
for (const reports of reportByPerson.values()) {
  reports.Inicio.sort((a, b) => new Date(a.email_ts) - new Date(b.email_ts));
  reports.Cierre.sort((a, b) => new Date(a.email_ts) - new Date(b.email_ts));
}

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Resumen");
const register = workbook.worksheets.add("Registro");
const matrix = workbook.worksheets.add("Matriz mensual");
const people = workbook.worksheets.add("Personal");

const colors = {
  navy: "#17365D",
  blue: "#2F75B5",
  paleBlue: "#D9EAF7",
  green: "#70AD47",
  paleGreen: "#E2F0D9",
  amber: "#FFC000",
  paleAmber: "#FFF2CC",
  red: "#C00000",
  paleRed: "#FCE4D6",
  gray: "#F2F2F2",
  border: "#D9E2F3",
  white: "#FFFFFF",
  text: "#1F2937",
};

function styleTitle(sheet, rangeAddress, title) {
  const range = sheet.getRange(rangeAddress);
  range.merge();
  range.values = [[title]];
  range.format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  range.format.rowHeight = 30;
}

function styleHeader(range) {
  range.format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: colors.border },
  };
  range.format.rowHeight = 30;
}

function excelColumn(columnNumber) {
  let value = columnNumber;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function localTimeValue(isoTimestamp) {
  const match = /T(\d{2}):(\d{2})/.exec(isoTimestamp ?? "");
  if (!match) return null;
  return (Number(match[1]) * 60 + Number(match[2])) / 1440;
}

for (const sheet of [summary, register, matrix, people]) {
  sheet.showGridLines = false;
}

// Personal
styleTitle(people, "A1:C1", "Personal sujeto a control de jornada");
people.getRange("A2:C2").merge();
people.getRange("A2").values = [["Lista base utilizada por las revisiones programadas"]];
people.getRange("A2").format = { fill: colors.paleBlue, font: { italic: true, color: colors.text } };
people.getRange("A4:C4").values = [["Nombre", "Correo", "Grupo"]];
styleHeader(people.getRange("A4:C4"));
people.getRange("A5:C20").values = seed.roster.map(([name, email, group]) => [name, email, group]);
people.getRange("A5:C20").format.borders = { preset: "inside", style: "thin", color: colors.border };
people.tables.add("A4:C20", true, "PersonalTable").style = "TableStyleMedium2";
people.freezePanes.freezeRows(4);
people.getRange("A1:A20").format.columnWidth = 36;
people.getRange("B1:B20").format.columnWidth = 34;
people.getRange("C1:C20").format.columnWidth = 24;

// Registro
styleTitle(register, "A1:N1", "Registro acumulativo de inicios y cierres");
register.getRange("A2:N2").merge();
register.getRange("A2").values = [[
  "Una fila por persona y fecha. La fecha corresponde a la recepción del mensaje en Gmail (America/Bogota)."
]];
register.getRange("A2").format = { fill: colors.paleBlue, font: { italic: true, color: colors.text }, wrapText: true };
const registerHeaders = [
  "Fecha", "Persona", "Correo", "Grupo", "Inicio", "Hora inicio",
  "Evidencia inicio", "ID mensaje inicio", "Cierre", "Hora cierre",
  "Evidencia cierre", "ID mensaje cierre", "Estado", "Observaciones"
];
register.getRange("A4:N4").values = [registerHeaders];
styleHeader(register.getRange("A4:N4"));

const dateValue = new Date(seed.date + "T00:00:00-05:00");
const registerRows = seed.roster.map(([name, email, group]) => {
  const reports = reportByPerson.get(email) ?? { Inicio: [], Cierre: [] };
  const start = reports.Inicio[0] ?? null;
  const close = reports.Cierre.at(-1) ?? null;
  const notes = [];
  if (start && !start.visualEvidence) notes.push("Inicio sin evidencia visual");
  if (close && !close.visualEvidence) notes.push("Cierre sin evidencia visual");
  if (start && /21[/-]07[/-]2026/.test(start.subject ?? "")) notes.push("El asunto del inicio menciona 21/07/2026; se registró por fecha de recepción");
  return [
    dateValue,
    name,
    email,
    group,
    start ? "Sí" : "No",
    start ? localTimeValue(start.email_ts) : null,
    start ? (start.visualEvidence ? "Sí" : "No") : "Sin reporte",
    start?.id ?? "",
    close ? "Sí" : "No",
    close ? localTimeValue(close.email_ts) : null,
    close ? (close.visualEvidence ? "Sí" : "No") : "Sin reporte",
    close?.id ?? "",
    null,
    notes.join("; "),
  ];
});
register.getRange("A5:N20").values = registerRows;
register.getRange("M5").formulas = [[
  '=IF(AND(E5="Sí",I5="Sí"),"Completo",IF(AND(E5="Sí",I5<>"Sí"),"Falta cierre",IF(AND(E5<>"Sí",I5="Sí"),"Falta inicio","Sin reportes")))'
]];
register.getRange("M5:M20").fillDown();
register.getRange("A5:A20").format.numberFormat = "yyyy-mm-dd";
register.getRange("F5:F20").format.numberFormat = "hh:mm";
register.getRange("J5:J20").format.numberFormat = "hh:mm";
register.getRange("A5:N20").format.borders = { preset: "inside", style: "thin", color: colors.border };
register.getRange("A5:N20").format.verticalAlignment = "center";
register.getRange("N5:N20").format.wrapText = true;
register.tables.add("A4:N20", true, "RegistroTable").style = "TableStyleMedium2";
register.getRange("M5:M20").conditionalFormats.add("containsText", {
  text: "Completo",
  format: { fill: colors.paleGreen, font: { color: "#385723", bold: true } },
});
register.getRange("M5:M20").conditionalFormats.add("containsText", {
  text: "Falta",
  format: { fill: colors.paleAmber, font: { color: "#7F6000", bold: true } },
});
register.getRange("M5:M20").conditionalFormats.add("containsText", {
  text: "Sin reportes",
  format: { fill: colors.paleRed, font: { color: colors.red, bold: true } },
});
register.freezePanes.freezeRows(4);
register.freezePanes.freezeColumns(2);
const registerWidths = [13, 34, 32, 22, 10, 12, 16, 22, 10, 12, 16, 22, 16, 46];
registerWidths.forEach((width, index) => register.getRangeByIndexes(0, index, 20, 1).format.columnWidth = width);
register.getRange("H1:H20").format.font = { color: "#6B7280", size: 8 };
register.getRange("L1:L20").format.font = { color: "#6B7280", size: 8 };

// Matriz mensual
styleTitle(matrix, "A1:AG1", "Matriz mensual de control de jornada");
matrix.getRange("A2:AG2").merge();
matrix.getRange("A2").values = [["Julio de 2026 · I = inicio · C = cierre · Sin I/C = reporte faltante"]];
matrix.getRange("A2").format = { fill: colors.paleBlue, font: { italic: true, color: colors.text } };
matrix.getRange("A4:B4").values = [["Persona", "Grupo"]];
const dayHeaders = Array.from({ length: 31 }, (_, i) => new Date("2026-07-" + String(i + 1).padStart(2, "0") + "T00:00:00-05:00"));
matrix.getRange("C4:AG4").values = [dayHeaders];
styleHeader(matrix.getRange("A4:AG4"));
matrix.getRange("C4:AG4").format.numberFormat = "dd";
matrix.getRange("A5:B20").values = seed.roster.map(([name, , group]) => [name, group]);

const matrixFormulas = seed.roster.map(([, email], rowIndex) => {
  const personalRow = rowIndex + 5;
  return dayHeaders.map((_, dayIndex) => {
    const column = excelColumn(3 + dayIndex);
    const headerRef = column + "$4";
    const emailRef = "'Personal'!$B$" + personalRow;
    return '=IF(COUNTIFS(\'Registro\'!$A$5:$A$500,' + headerRef + ',\'Registro\'!$C$5:$C$500,' + emailRef + ')=0,"",'
      + 'IF(COUNTIFS(\'Registro\'!$A$5:$A$500,' + headerRef + ',\'Registro\'!$C$5:$C$500,' + emailRef + ',\'Registro\'!$E$5:$E$500,"Sí")>0,'
      + '"I "&TEXT(SUMIFS(\'Registro\'!$F$5:$F$500,\'Registro\'!$A$5:$A$500,' + headerRef + ',\'Registro\'!$C$5:$C$500,' + emailRef + '),"hh:mm"),"Sin I")'
      + '&" | "&'
      + 'IF(COUNTIFS(\'Registro\'!$A$5:$A$500,' + headerRef + ',\'Registro\'!$C$5:$C$500,' + emailRef + ',\'Registro\'!$I$5:$I$500,"Sí")>0,'
      + '"C "&TEXT(SUMIFS(\'Registro\'!$J$5:$J$500,\'Registro\'!$A$5:$A$500,' + headerRef + ',\'Registro\'!$C$5:$C$500,' + emailRef + '),"hh:mm"),"Sin C"))';
  });
});
matrix.getRange("C5:AG20").formulas = matrixFormulas;
matrix.getRange("A5:AG20").format.borders = { preset: "inside", style: "thin", color: colors.border };
matrix.getRange("C5:AG20").format = {
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  font: { size: 8 },
};
matrix.getRange("C5:AG20").conditionalFormats.add("containsText", {
  text: "Sin",
  format: { fill: colors.paleAmber, font: { color: "#7F6000", bold: true } },
});
matrix.getRange("C5:AG20").conditionalFormats.addCustom(
  '=AND(C5<>"",ISERROR(SEARCH("Sin",C5)))',
  { fill: colors.paleGreen, font: { color: "#385723", bold: true } },
);
matrix.freezePanes.freezeRows(4);
matrix.freezePanes.freezeColumns(2);
matrix.getRange("A1:A20").format.columnWidth = 34;
matrix.getRange("B1:B20").format.columnWidth = 21;
matrix.getRange("C1:AG20").format.columnWidth = 16;
matrix.getRange("A5:AG20").format.rowHeight = 30;

// Resumen
styleTitle(summary, "A1:H1", "Control diario de jornada");
summary.getRange("A2:H2").merge();
summary.getRange("A2").values = [["Actualizado con los mensajes recibidos el 22 de julio de 2026"]];
summary.getRange("A2").format = { fill: colors.paleBlue, font: { italic: true, color: colors.text } };
summary.getRange("A4:B4").merge();
summary.getRange("C4:D4").merge();
summary.getRange("E4:F4").merge();
summary.getRange("G4:H4").merge();
summary.getRange("A4").values = [["Inicios recibidos"]];
summary.getRange("C4").values = [["Cierres recibidos"]];
summary.getRange("E4").values = [["Registros completos"]];
summary.getRange("G4").values = [["Personas con pendientes"]];
for (const address of ["A4:B4", "C4:D4", "E4:F4", "G4:H4"]) {
  summary.getRange(address).format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
  };
}
summary.getRange("A5:B6").merge();
summary.getRange("C5:D6").merge();
summary.getRange("E5:F6").merge();
summary.getRange("G5:H6").merge();
summary.getRange("A5").formulas = [['=COUNTIFS(\'Registro\'!$A$5:$A$500,DATE(2026,7,22),\'Registro\'!$E$5:$E$500,"Sí")']];
summary.getRange("C5").formulas = [['=COUNTIFS(\'Registro\'!$A$5:$A$500,DATE(2026,7,22),\'Registro\'!$I$5:$I$500,"Sí")']];
summary.getRange("E5").formulas = [['=COUNTIFS(\'Registro\'!$A$5:$A$500,DATE(2026,7,22),\'Registro\'!$M$5:$M$500,"Completo")']];
summary.getRange("G5").formulas = [['=COUNTIFS(\'Registro\'!$A$5:$A$500,DATE(2026,7,22),\'Registro\'!$M$5:$M$500,"<>Completo")']];
for (const address of ["A5:B6", "C5:D6", "E5:F6", "G5:H6"]) {
  summary.getRange(address).format = {
    fill: colors.gray,
    font: { bold: true, color: colors.navy, size: 22 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: colors.border },
  };
}
summary.getRange("A8:H8").values = [[
  "Persona", "Grupo", "Inicio", "Hora inicio", "Cierre", "Hora cierre", "Estado", "Observaciones"
]];
styleHeader(summary.getRange("A8:H8"));
summary.getRange("A9:H24").formulas = seed.roster.map((_, index) => {
  const registerRow = index + 5;
  return [
    "='Registro'!B" + registerRow,
    "='Registro'!D" + registerRow,
    "='Registro'!E" + registerRow,
    '=IF(\'Registro\'!F' + registerRow + '="","",\'Registro\'!F' + registerRow + ')',
    "='Registro'!I" + registerRow,
    '=IF(\'Registro\'!J' + registerRow + '="","",\'Registro\'!J' + registerRow + ')',
    "='Registro'!M" + registerRow,
    '=IF(\'Registro\'!N' + registerRow + '="","",\'Registro\'!N' + registerRow + ')',
  ];
});
summary.getRange("D9:D24").format.numberFormat = "hh:mm";
summary.getRange("F9:F24").format.numberFormat = "hh:mm";
summary.getRange("A9:H24").format.borders = { preset: "inside", style: "thin", color: colors.border };
summary.getRange("H9:H24").format.wrapText = true;
summary.getRange("G9:G24").conditionalFormats.add("containsText", {
  text: "Completo",
  format: { fill: colors.paleGreen, font: { color: "#385723", bold: true } },
});
summary.getRange("G9:G24").conditionalFormats.add("containsText", {
  text: "Falta",
  format: { fill: colors.paleAmber, font: { color: "#7F6000", bold: true } },
});
summary.getRange("G9:G24").conditionalFormats.add("containsText", {
  text: "Sin reportes",
  format: { fill: colors.paleRed, font: { color: colors.red, bold: true } },
});
summary.freezePanes.freezeRows(8);
const summaryWidths = [34, 22, 11, 12, 11, 12, 18, 48];
summaryWidths.forEach((width, index) => summary.getRangeByIndexes(0, index, 100, 1).format.columnWidth = width);

await fs.mkdir(previewDir, { recursive: true });

const inspection = await workbook.inspect({
  kind: "table",
  range: "Resumen!A1:H24",
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 10,
  maxChars: 10000,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 3000,
});

for (const sheetName of ["Resumen", "Registro", "Matriz mensual", "Personal"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  const safeName = sheetName.toLowerCase().replace(/\s+/g, "_");
  await fs.writeFile(path.join(previewDir, safeName + ".png"), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(JSON.stringify({
  outputPath,
  previewDir,
  inspection: inspection.ndjson,
  errors: errors.ndjson,
}, null, 2));
