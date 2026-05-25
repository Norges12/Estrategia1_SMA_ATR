function main(workbook: ExcelScript.Workbook) {
  const meses = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];

  for (const mes of meses) {
    let sheet = workbook.getWorksheet(mes);
    if (!sheet) {
      sheet = workbook.addWorksheet(mes);
    }

    clearSheet(sheet);
    buildMonthSheet(sheet, mes);
  }

  buildResumenAnual(workbook, meses);
  deleteDefaultEmptySheet(workbook);
}

function clearSheet(sheet: ExcelScript.Worksheet): void {
  const used = sheet.getUsedRange();
  if (used) {
    used.clear(ExcelScript.ClearApplyTo.all);
  }

  for (const chart of sheet.getCharts()) {
    chart.delete();
  }

  sheet.setShowGridlines(false);
}

function buildMonthSheet(sheet: ExcelScript.Worksheet, mes: string): void {
  configureColumns(sheet);
  configureRows(sheet);
  setBaseFont(sheet);

  buildSidebar(sheet, mes);
  buildTopCards(sheet);
  buildFacturasBlock(sheet);
  buildVariablesBlock(sheet);
  buildAhorrosBlock(sheet);
  buildSeguimientoBlock(sheet);
  buildDeudaBlock(sheet);
  buildInversionesBlock(sheet);
  applyNumberFormats(sheet);
  applyAlignments(sheet);
  addCharts(sheet);
}

function configureColumns(sheet: ExcelScript.Worksheet): void {
  // OJO: en Office Scripts el ancho va en PUNTOS (no en "ancho de caracteres").
  // 1 punto ~ 1.33 px. Una columna normal de Excel ~ 48 pt.
  const widths: { [key: string]: number } = {
    "A:A": 24,     // margen izquierdo

    "B:B": 120,    // concepto
    "C:C": 80,
    "D:D": 80,
    "E:E": 14,     // separador

    "F:F": 28,     // OK / check
    "G:G": 150,    // concepto
    "H:H": 80,
    "I:I": 80,
    "J:J": 95,     // tipo
    "K:K": 70,     // fecha

    "L:L": 28,
    "M:M": 150,
    "N:N": 80,
    "O:O": 80,
    "P:P": 100,    // tipo
    "Q:Q": 14,

    "R:R": 28,
    "S:S": 150,
    "T:T": 80,
    "U:U": 110,    // categoria
    "V:V": 75,     // fecha
    "W:W": 14,

    "X:X": 28,
    "Y:Y": 80,
    "Z:Z": 80,
    "AA:AA": 80,
    "AB:AB": 80,
    "AC:AC": 14,

    "AD:AD": 28,
    "AE:AE": 150,  // concepto
    "AF:AF": 80,
    "AG:AG": 80,
    "AH:AH": 100,  // tipo
    "AI:AI": 75    // fecha
  };

  for (const address of Object.keys(widths)) {
    sheet.getRange(address).getFormat().setColumnWidth(widths[address]);
  }
}

function configureRows(sheet: ExcelScript.Worksheet): void {
  for (let row = 1; row <= 45; row++) {
    let height = 16;                              // filas de datos

    if (row >= 2 && row <= 10) height = 22;       // zona de graficos / titulo
    if (row === 7) height = 20;                   // VISTA GENERAL header
    if (row === 12 || row === 31) height = 20;    // headers de bloques
    if (row >= 11 && row <= 16) height = 22;      // tarjetas KPI

    sheet.getRange(`${row}:${row}`).getFormat().setRowHeight(height);
  }
}

function setBaseFont(sheet: ExcelScript.Worksheet): void {
  const all = sheet.getRange("A1:AI45");
  all.getFormat().getFont().setName("Aptos");
  all.getFormat().getFont().setSize(9);
  all.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function buildSidebar(sheet: ExcelScript.Worksheet, mes: string): void {
  mergeCell(sheet, "A2:D4", mes, 20, true, "000000", "FFFFFF");
  mergeCell(sheet, "A5:D5", "-Dashboard-", 10, false, "444444", "FFFFFF");

  setSectionHeader(sheet, "A7:D7", "VISTA GENERAL", "B4C7E7", "000000");
  sheet.getRange("A8:B9").setValues([
    ["Año", new Date().getFullYear()],
    ["Mes", mes]
  ]);
  sheet.getRange("C8:D9").setValues([
    ["Moneda", "USD"],
    ["Estado", "Activo"]
  ]);
  applyBorders(sheet.getRange("A8:D9"), "D9D9D9");

  createKpiCard(sheet, "A11:B13", "INGRESO TOTAL", "=SUM(D20:D23)", "FFFFFF");
  createKpiCard(sheet, "C11:D13", "TOTAL GASTADO", "=SUM(I14:I35)+SUM(O14:O28)+SUM(AG14:AG28)+SUM(AG33:AG38)", "FFFFFF");
  createKpiCard(sheet, "A14:B16", "PRESUPUESTO + AHORRO", "=SUM(H14:H35)+SUM(N14:N28)+SUM(AF14:AF28)+SUM(AF33:AF38)", "FFFFFF");
  createKpiCard(sheet, "C14:D16", "TOTAL AHORRADO", "=SUM(O33:O38)+SUM(AG33:AG38)", "FFFFFF");

  buildIngresosSidebar(sheet);
  buildBudget503020(sheet);
  buildSidebarComparison(sheet);
}

function buildIngresosSidebar(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "A18:D18", "INGRESOS", "D9EAD3", "000000");
  sheet.getRange("A19:D19").setValues([["CONCEPTO", "TIPO", "PRESUPUESTO", "ACTUAL"]]);
  styleTableHeader(sheet.getRange("A19:D19"), "EAF4E3", "000000");

  sheet.getRange("A20:D23").setValues([
    ["Sueldo", "Fijo", 1000, 1400],
    ["Freelance", "Variable", 400, 0],
    ["Inversiones", "Variable", 0, 0],
    ["Otros", "Variable", 0, 0]
  ]);

  sheet.getRange("A24").setValue("TOTAL");
  sheet.getRange("C24").setFormula("=SUM(C20:C23)");
  sheet.getRange("D24").setFormula("=SUM(D20:D23)");
  styleTotalRow(sheet.getRange("A24:D24"), "D9EAD3");
  applyBorders(sheet.getRange("A19:D24"), "D9D9D9");
}

function buildBudget503020(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "A26:D26", "PRESUPUESTO 50/30/20", "FCE5CD", "000000");
  sheet.getRange("A27:D27").setValues([["CATEGORÍA", "%", "PRESUPUESTO", "ACTUAL"]]);
  styleTableHeader(sheet.getRange("A27:D27"), "FDF2E9", "000000");

  sheet.getRange("A28:B30").setValues([
    ["Necesidades", 0.5],
    ["Deseos", 0.3],
    ["Ahorros", 0.2]
  ]);

  sheet.getRange("C28").setFormula("=D24*B28");
  sheet.getRange("C29").setFormula("=D24*B29");
  sheet.getRange("C30").setFormula("=D24*B30");

  sheet.getRange("D28").setFormula("=SUM(I14:I35)+SUM(AG14:AG28)");
  sheet.getRange("D29").setFormula("=SUM(O14:O28)");
  sheet.getRange("D30").setFormula("=SUM(O33:O38)+SUM(AG33:AG38)");

  sheet.getRange("A31").setValue("TOTAL");
  sheet.getRange("B31").setValue(1);
  sheet.getRange("C31").setFormula("=SUM(C28:C30)");
  sheet.getRange("D31").setFormula("=SUM(D28:D30)");

  styleTotalRow(sheet.getRange("A31:D31"), "FCE5CD");
  applyBorders(sheet.getRange("A27:D31"), "D9D9D9");
}

function buildSidebarComparison(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "A33:D33", "PRESUPUESTO VS ACTUAL", "D9D2E9", "000000");
  sheet.getRange("A34:D34").setValues([["RUBRO", "PRESUPUESTO", "ACTUAL", "DIF."]]);
  styleTableHeader(sheet.getRange("A34:D34"), "EEEAF7", "000000");

  sheet.getRange("A35:A40").setValues([
    ["Ingresos"],
    ["Gastos fijos"],
    ["Variables"],
    ["Ahorros"],
    ["Deuda"],
    ["Restante"]
  ]);

  sheet.getRange("B35").setFormula("=SUM(C20:C23)");
  sheet.getRange("C35").setFormula("=SUM(D20:D23)");

  sheet.getRange("B36").setFormula("=SUM(H14:H35)");
  sheet.getRange("C36").setFormula("=SUM(I14:I35)");

  sheet.getRange("B37").setFormula("=SUM(N14:N28)");
  sheet.getRange("C37").setFormula("=SUM(O14:O28)");

  sheet.getRange("B38").setFormula("=SUM(N33:N38)");
  sheet.getRange("C38").setFormula("=SUM(O33:O38)");

  sheet.getRange("B39").setFormula("=SUM(AF14:AF28)");
  sheet.getRange("C39").setFormula("=SUM(AG14:AG28)");

  sheet.getRange("B40").setFormula("=B35-(B36+B37+B38+B39)");
  sheet.getRange("C40").setFormula("=C35-(C36+C37+C38+C39)");

  for (let row = 35; row <= 40; row++) {
    sheet.getRange(`D${row}`).setFormula(`=C${row}-B${row}`);
  }

  applyBorders(sheet.getRange("A34:D40"), "D9D9D9");
}

function buildTopCards(sheet: ExcelScript.Worksheet): void {
  createTopCard(sheet, "F2:K10", "50/30/20 PRESUPUESTO");
  createTopCard(sheet, "L2:P10", "INGRESOS VS GASTOS");
  createTopCard(sheet, "R2:V10", "GASTOS");
  createTopCard(sheet, "X2:AB10", "DINERO RESTANTE");
  createTopCard(sheet, "AD2:AI10", "Presupuesto vs Actual");
}

function buildFacturasBlock(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "F12:K12", "FACTURAS / GASTOS FIJOS", "F4CCCC", "000000");
  sheet.getRange("F13:K13").setValues([["OK", "CONCEPTO", "PRESUPUESTO", "ACTUAL", "TIPO", "FECHA"]]);
  styleTableHeader(sheet.getRange("F13:K13"), "FCE5CD", "000000");

  const items = [
    ["☐", "Renta", 400, 400, "Necesidades", "15-abr"],
    ["☐", "Agua", 40, 35, "Necesidades", "18-abr"],
    ["☐", "Internet", 50, 50, "Necesidades", "24-abr"],
    ["☐", "Luz", 60, 40, "Necesidades", "19-abr"],
    ["☐", "Teléfono", 0, 0, "", ""],
    ["☐", "Gas", 0, 0, "", ""],
    ["☐", "Seguro", 0, 0, "", ""],
    ["☐", "Suscripciones", 0, 0, "", ""],
    ["☐", "Gym", 0, 0, "", ""],
    ["☐", "Streaming", 0, 0, "", ""],
    ["☐", "Crédito", 0, 0, "", ""],
    ["☐", "Cuota auto", 0, 0, "", ""],
    ["☐", "Salud", 0, 0, "", ""],
    ["☐", "Impuestos", 0, 0, "", ""],
    ["☐", "Colegio", 0, 0, "", ""],
    ["☐", "Comunidad", 0, 0, "", ""],
    ["☐", "Parking", 0, 0, "", ""],
    ["☐", "Otros", 0, 0, "", ""],
    ["☐", "", 0, 0, "", ""],
    ["☐", "", 0, 0, "", ""],
    ["☐", "", 0, 0, "", ""],
    ["☐", "", 0, 0, "", ""]
  ];

  for (let i = 0; i < items.length; i++) {
    const row = 14 + i;
    sheet.getRange(`F${row}:K${row}`).setValues([items[i]]);
  }

  sheet.getRange("F36").setValue("TOTAL");
  sheet.getRange("H36").setFormula("=SUM(H14:H35)");
  sheet.getRange("I36").setFormula("=SUM(I14:I35)");
  styleTotalRow(sheet.getRange("F36:K36"), "F4CCCC");
  applyBorders(sheet.getRange("F13:K36"), "D9D9D9");
}

function buildVariablesBlock(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "L12:P12", "GASTOS VARIABLES", "D9D2E9", "000000");
  sheet.getRange("L13:P13").setValues([["OK", "CONCEPTO", "PRESUPUESTO", "ACTUAL", "TIPO"]]);
  styleTableHeader(sheet.getRange("L13:P13"), "EEEAF7", "000000");

  const items = [
    ["☐", "Mercado", 150, 85, "Necesidades"],
    ["☐", "Restaurantes", 60, 0, "Deseos"],
    ["☐", "Entretenimiento", 40, 0, "Deseos"],
    ["☐", "Salud", 30, 10, "Necesidades"],
    ["☐", "Cuidado personal", 30, 0, "Necesidades"],
    ["☐", "Ropa", 30, 0, "Deseos"],
    ["☐", "Educación", 50, 0, "Necesidades"],
    ["☐", "Vacaciones", 0, 0, "Deseos"],
    ["☐", "Mascotas", 30, 20, "Necesidades"],
    ["☐", "Regalos", 20, 0, "Deseos"],
    ["☐", "Medicamentos", 20, 0, "Necesidades"],
    ["☐", "Otros", 0, 0, "Deseos"],
    ["☐", "", 0, 0, ""],
    ["☐", "", 0, 0, ""],
    ["☐", "", 0, 0, ""]
  ];

  for (let i = 0; i < items.length; i++) {
    const row = 14 + i;
    sheet.getRange(`L${row}:P${row}`).setValues([items[i]]);
  }

  sheet.getRange("L29").setValue("TOTAL");
  sheet.getRange("N29").setFormula("=SUM(N14:N28)");
  sheet.getRange("O29").setFormula("=SUM(O14:O28)");
  styleTotalRow(sheet.getRange("L29:P29"), "D9D2E9");
  applyBorders(sheet.getRange("L13:P29"), "D9D9D9");
}

function buildAhorrosBlock(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "L31:P31", "AHORROS", "F9CB9C", "000000");
  sheet.getRange("L32:P32").setValues([["OK", "CONCEPTO", "PRESUPUESTO", "ACTUAL", "TIPO"]]);
  styleTableHeader(sheet.getRange("L32:P32"), "FCE5CD", "000000");

  const items = [
    ["☐", "Viaje a Europa", 400, 400, "Ahorro"],
    ["☐", "Retiro", 200, 150, "Ahorro"],
    ["☐", "Fondo emergencia", 100, 0, "Ahorro"],
    ["☐", "Casa", 0, 0, "Ahorro"],
    ["☐", "Educación", 0, 0, "Ahorro"],
    ["☐", "Otros", 0, 0, "Ahorro"]
  ];

  for (let i = 0; i < items.length; i++) {
    const row = 33 + i;
    sheet.getRange(`L${row}:P${row}`).setValues([items[i]]);
  }

  sheet.getRange("L39").setValue("TOTAL");
  sheet.getRange("N39").setFormula("=SUM(N33:N38)");
  sheet.getRange("O39").setFormula("=SUM(O33:O38)");
  styleTotalRow(sheet.getRange("L39:P39"), "F9CB9C");
  applyBorders(sheet.getRange("L32:P39"), "D9D9D9");
}

function buildSeguimientoBlock(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "R12:V12", "SEGUIMIENTO DE GASTOS", "D9D2E9", "000000");
  sheet.getRange("R13:V13").setValues([["OK", "CONCEPTO", "MONTO", "CATEGORÍA", "FECHA"]]);
  styleTableHeader(sheet.getRange("R13:V13"), "EEEAF7", "000000");

  const items = [
    ["☐", "Carne", 80, "Mercado", "16-abr"],
    ["☐", "Oliva", 20, "Mercado", "20-abr"],
    ["☐", "Consulta médica", 30, "Salud", "21-abr"],
    ["☐", "Comida gatos", 40, "Mascotas", "27-abr"]
  ];

  for (let i = 0; i < 26; i++) {
    const row = 14 + i;
    sheet.getRange(`R${row}`).setValue("☐");

    if (i < items.length) {
      sheet.getRange(`R${row}:V${row}`).setValues([items[i]]);
    } else {
      sheet.getRange(`R${row}`).setValue("☐");
    }
  }

  sheet.getRange("R40").setValue("TOTAL");
  sheet.getRange("T40").setFormula("=SUM(T14:T39)");
  styleTotalRow(sheet.getRange("R40:V40"), "D9D2E9");
  applyBorders(sheet.getRange("R13:V40"), "D9D9D9");
}

function buildDeudaBlock(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "AD12:AI12", "PAGO DE DEUDA", "9FC5E8", "000000");
  sheet.getRange("AD13:AI13").setValues([["OK", "CONCEPTO", "PRESUPUESTO", "ACTUAL", "TIPO", "FECHA"]]);
  styleTableHeader(sheet.getRange("AD13:AI13"), "D9EAF7", "000000");

  const items = [
    ["☐", "Tarjeta de crédito", 0, 0, "Deuda", ""],
    ["☐", "Auto", 0, 0, "Deuda", ""],
    ["☐", "Préstamo personal", 0, 0, "Deuda", ""],
    ["☐", "Hipoteca", 0, 0, "Deuda", ""]
  ];

  for (let i = 0; i < 15; i++) {
    const row = 14 + i;

    if (i < items.length) {
      sheet.getRange(`AD${row}:AI${row}`).setValues([items[i]]);
    } else {
      sheet.getRange(`AD${row}`).setValue("☐");
    }
  }

  sheet.getRange("AD29").setValue("TOTAL");
  sheet.getRange("AF29").setFormula("=SUM(AF14:AF28)");
  sheet.getRange("AG29").setFormula("=SUM(AG14:AG28)");
  styleTotalRow(sheet.getRange("AD29:AI29"), "9FC5E8");
  applyBorders(sheet.getRange("AD13:AI29"), "D9D9D9");
}

function buildInversionesBlock(sheet: ExcelScript.Worksheet): void {
  setSectionHeader(sheet, "AD31:AI31", "INVERSIONES / DONACIONES", "D9EAD3", "000000");
  sheet.getRange("AD32:AI32").setValues([["OK", "CONCEPTO", "PRESUPUESTO", "ACTUAL", "TIPO", "FECHA"]]);
  styleTableHeader(sheet.getRange("AD32:AI32"), "EAF4E3", "000000");

  const items = [
    ["☐", "ETF", 0, 0, "Inversión", ""],
    ["☐", "Acciones", 0, 0, "Inversión", ""],
    ["☐", "Crypto", 0, 0, "Inversión", ""],
    ["☐", "Donación", 0, 0, "Donación", ""],
    ["☐", "ONG", 0, 0, "Donación", ""],
    ["☐", "Otros", 0, 0, "Inversión", ""]
  ];

  for (let i = 0; i < items.length; i++) {
    const row = 33 + i;
    sheet.getRange(`AD${row}:AI${row}`).setValues([items[i]]);
  }

  sheet.getRange("AD39").setValue("TOTAL");
  sheet.getRange("AF39").setFormula("=SUM(AF33:AF38)");
  sheet.getRange("AG39").setFormula("=SUM(AG33:AG38)");
  styleTotalRow(sheet.getRange("AD39:AI39"), "D9EAD3");
  applyBorders(sheet.getRange("AD32:AI39"), "D9D9D9");
}

function applyNumberFormats(sheet: ExcelScript.Worksheet): void {
  const moneyRanges = [
    "C20:D24",
    "C28:D31",
    "B35:D40",
    "H14:I36",
    "N14:O29",
    "N33:O39",
    "T14:T40",
    "AF14:AG29",
    "AF33:AG39"
  ];

  for (const address of moneyRanges) {
    sheet.getRange(address).setNumberFormat("$#,##0.00");
  }

  sheet.getRange("B28:B31").setNumberFormat("0%");
}

function applyAlignments(sheet: ExcelScript.Worksheet): void {
  const centered = [
    "A8:D9",
    "A19:D24",
    "A27:D31",
    "A34:D40",
    "F13:K36",
    "L13:P29",
    "L32:P39",
    "R13:V40",
    "AD13:AI29",
    "AD32:AI39"
  ];

  for (const address of centered) {
    sheet.getRange(address).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }

  const leftRanges = [
    "A20:A23",
    "A28:A31",
    "A35:A40",
    "G14:G35",
    "J14:J35",
    "M14:M28",
    "P14:P28",
    "M33:M38",
    "P33:P38",
    "S14:S39",
    "U14:U39",
    "AE14:AE28",
    "AH14:AH28",
    "AE33:AE38",
    "AH33:AH38"
  ];

  for (const address of leftRanges) {
    const format = sheet.getRange(address).getFormat();
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    format.setWrapText(false);
  }

  const headerRanges = [
    "A19:D19",
    "A27:D27",
    "A34:D34",
    "F13:K13",
    "L13:P13",
    "L32:P32",
    "R13:V13",
    "AD13:AI13",
    "AD32:AI32"
  ];

  for (const address of headerRanges) {
    const format = sheet.getRange(address).getFormat();
    format.setWrapText(true);
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }
}

function addCharts(sheet: ExcelScript.Worksheet): void {
  // Datos auxiliares para los graficos de dona/barras (zona "oculta" en blanco)
  sheet.getRange("X34:Y35").setValues([
    ["Gastado", 1],
    ["Restante", 1]
  ]);
  sheet.getRange("X37:Z40").setValues([
    ["Ingresos", 0, 0],
    ["Gastos", 0, 0],
    ["Ahorros", 0, 0],
    ["Deuda", 0, 0]
  ]);
  sheet.getRange("Y34").setFormula("=C12");
  sheet.getRange("Y35").setFormula("=MAX(A12-C12,0)");

  sheet.getRange("Y37").setFormula("=B35");
  sheet.getRange("Z37").setFormula("=C35");
  sheet.getRange("Y38").setFormula("=B36");
  sheet.getRange("Z38").setFormula("=C36");
  sheet.getRange("Y39").setFormula("=B38");
  sheet.getRange("Z39").setFormula("=C38");
  sheet.getRange("Y40").setFormula("=B39");
  sheet.getRange("Z40").setFormula("=C39");

  const chart1 = sheet.addChart(ExcelScript.ChartType.doughnut, sheet.getRange("A28:A30").getResizedRange(0, 2));
  chart1.setPosition("F3", "K10");
  chart1.getTitle().setText("50/30/20 PRESUPUESTO");

  const chart2 = sheet.addChart(ExcelScript.ChartType.columnClustered, sheet.getRange("A35:C36"));
  chart2.setPosition("L3", "P10");
  chart2.getTitle().setText("INGRESOS VS GASTOS");

  const chart3 = sheet.addChart(ExcelScript.ChartType.doughnut, sheet.getRange("A35:C39"));
  chart3.setPosition("R3", "V10");
  chart3.getTitle().setText("GASTOS");

  const chart4 = sheet.addChart(ExcelScript.ChartType.doughnut, sheet.getRange("X34:Y35"));
  chart4.setPosition("X3", "AB10");
  chart4.getTitle().setText("DINERO RESTANTE");

  const chart5 = sheet.addChart(ExcelScript.ChartType.barClustered, sheet.getRange("X37:Z40"));
  chart5.setPosition("AD3", "AI10");
  chart5.getTitle().setText("Presupuesto vs Actual");

  // Pintamos en blanco los datos auxiliares para que no se vean
  sheet.getRange("X34:Z40").getFormat().getFont().setColor("FFFFFF");
}

function createTopCard(sheet: ExcelScript.Worksheet, address: string, title: string): void {
  const range = sheet.getRange(address);
  range.getFormat().getFill().setColor("FFFFFF");
  applyBorders(range, "D9D9D9");

  const titleRow = sheet.getRangeByIndexes(
    range.getRowIndex(),
    range.getColumnIndex(),
    1,
    range.getColumnCount()
  );

  titleRow.merge(false);
  titleRow.setValue(title);
  titleRow.getFormat().getFont().setBold(true);
  titleRow.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
}

function createKpiCard(
  sheet: ExcelScript.Worksheet,
  address: string,
  title: string,
  formula: string,
  fillColor: string
): void {
  const range = sheet.getRange(address);
  range.getFormat().getFill().setColor(fillColor);
  applyBorders(range, "D9D9D9");

  const top = sheet.getRangeByIndexes(
    range.getRowIndex(),
    range.getColumnIndex(),
    1,
    range.getColumnCount()
  );
  const bottom = sheet.getRangeByIndexes(
    range.getRowIndex() + 1,
    range.getColumnIndex(),
    range.getRowCount() - 1,
    range.getColumnCount()
  );

  top.merge(false);
  bottom.merge(false);

  top.setValue(title);
  bottom.setFormula(formula);
  bottom.setNumberFormat("$#,##0.00");

  top.getFormat().getFont().setBold(true);
  top.getFormat().getFont().setSize(8);
  top.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  bottom.getFormat().getFont().setBold(true);
  bottom.getFormat().getFont().setSize(13);
  bottom.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
}

function setSectionHeader(
  sheet: ExcelScript.Worksheet,
  address: string,
  text: string,
  fillColor: string,
  fontColor: string
): void {
  const range = sheet.getRange(address);
  range.merge(false);
  range.setValue(text);

  const format = range.getFormat();
  format.getFont().setBold(true);
  format.getFont().setColor(fontColor);
  format.getFill().setColor(fillColor);
  format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  format.setWrapText(true);

  applyBorders(range, "D9D9D9");
}

function styleTableHeader(range: ExcelScript.Range, fillColor: string, fontColor: string): void {
  const format = range.getFormat();
  format.getFont().setBold(true);
  format.getFont().setColor(fontColor);
  format.getFill().setColor(fillColor);
  format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  format.setWrapText(true);
}

function styleTotalRow(range: ExcelScript.Range, fillColor: string): void {
  range.getFormat().getFont().setBold(true);
  range.getFormat().getFill().setColor(fillColor);
  applyBorders(range, "D9D9D9");
}

function mergeCell(
  sheet: ExcelScript.Worksheet,
  address: string,
  text: string,
  fontSize: number,
  bold: boolean,
  fontColor: string,
  fillColor: string
): void {
  const range = sheet.getRange(address);
  range.merge(false);
  range.setValue(text);
  range.getFormat().getFont().setSize(fontSize);
  range.getFormat().getFont().setBold(bold);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFill().setColor(fillColor);
  range.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function applyBorder(
  format: ExcelScript.RangeFormat,
  borderIndex: ExcelScript.BorderIndex,
  color: string
): void {
  const border = format.getRangeBorder(borderIndex);
  border.setStyle(ExcelScript.BorderLineStyle.continuous);
  border.setWeight(ExcelScript.BorderWeight.thin);
  border.setColor(color);
}

function applyBorders(range: ExcelScript.Range, color: string): void {
  const format = range.getFormat();
  applyBorder(format, ExcelScript.BorderIndex.edgeTop, color);
  applyBorder(format, ExcelScript.BorderIndex.edgeBottom, color);
  applyBorder(format, ExcelScript.BorderIndex.edgeLeft, color);
  applyBorder(format, ExcelScript.BorderIndex.edgeRight, color);
  applyBorder(format, ExcelScript.BorderIndex.insideHorizontal, color);
  applyBorder(format, ExcelScript.BorderIndex.insideVertical, color);
}

function buildResumenAnual(workbook: ExcelScript.Workbook, meses: string[]): void {
  let sheet = workbook.getWorksheet("Resumen Anual");
  if (!sheet) {
    sheet = workbook.addWorksheet("Resumen Anual");
  }

  clearSheet(sheet);

  sheet.getRange("A1:F1").merge(false);
  sheet.getRange("A1").setValue("RESUMEN ANUAL CONSOLIDADO");
  sheet.getRange("A1").getFormat().getFont().setBold(true);
  sheet.getRange("A1").getFormat().getFont().setSize(16);
  sheet.getRange("A1").getFormat().getFill().setColor("1F4E78");
  sheet.getRange("A1").getFormat().getFont().setColor("FFFFFF");
  sheet.getRange("A1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sheet.getRange("A3:F3").setValues([["MES", "INGRESOS", "GASTADO", "PRESUPUESTO+AHORRO", "AHORRADO", "RESTANTE"]]);
  styleTableHeader(sheet.getRange("A3:F3"), "D9EAF7", "000000");

  for (let i = 0; i < meses.length; i++) {
    const row = 4 + i;
    const mes = meses[i];

    sheet.getRange(`A${row}`).setValue(mes);
    sheet.getRange(`B${row}`).setFormula(`='${mes}'!A12`);
    sheet.getRange(`C${row}`).setFormula(`='${mes}'!C12`);
    sheet.getRange(`D${row}`).setFormula(`='${mes}'!A15`);
    sheet.getRange(`E${row}`).setFormula(`='${mes}'!C15`);
    sheet.getRange(`F${row}`).setFormula(`=B${row}-C${row}`);
  }

  sheet.getRange("A16").setValue("TOTAL");
  sheet.getRange("B16").setFormula("=SUM(B4:B15)");
  sheet.getRange("C16").setFormula("=SUM(C4:C15)");
  sheet.getRange("D16").setFormula("=SUM(D4:D15)");
  sheet.getRange("E16").setFormula("=SUM(E4:E15)");
  sheet.getRange("F16").setFormula("=SUM(F4:F15)");

  styleTotalRow(sheet.getRange("A16:F16"), "D9EAD3");
  applyBorders(sheet.getRange("A3:F16"), "D9D9D9");
  sheet.getRange("B4:F16").setNumberFormat("$#,##0.00");
  sheet.getRange("A:F").getFormat().autofitColumns();
}

function deleteDefaultEmptySheet(workbook: ExcelScript.Workbook): void {
  const defaultSheet = workbook.getWorksheet("Hoja1") ?? workbook.getWorksheet("Sheet1");
  if (!defaultSheet) {
    return;
  }

  const used = defaultSheet.getUsedRange();
  if (!used) {
    defaultSheet.delete();
    return;
  }

  const values = used.getValues();
  const hasContent = values.some(row => row.some(cell => cell !== "" && cell !== null));

  if (!hasContent) {
    defaultSheet.delete();
  }
}
