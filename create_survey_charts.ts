// Office Scripts: select Analysis or Analysis FR, then run this script.
// Creates a NEW chart sheet on every run. Does not change source data.
interface QuestionBlock {
  title: string;
  question: number;
  row: number;
  column: number;
  labels: string[];
  percentages: number[];
  base: number;
  kind: string;
}

function main(workbook: ExcelScript.Workbook) {
  const source = workbook.getActiveWorksheet();
  const used = source.getUsedRange(true);
  if (!used) throw new Error("Select your Analysis sheet before running the script.");
  const values = used.getValues();
  const blocks = findQuestions(values, used.getRowIndex(), used.getColumnIndex());
  if (blocks.length === 0) {
    throw new Error("No question tables found. Select Analysis or Analysis FR. Expected columns: Category/Catégorie, Institutions, Base (n), year percentage.");
  }
  const french = /\bFR\b/i.test(source.getName()) || values.some(row => row.some(v => normalize(String(v)) === "categorie"));
  const root = french ? "Charts FR Auto" : "Charts EN Auto";
  let name = root;
  let suffix = 2;
  while (workbook.getWorksheet(name)) { name = root + " " + suffix; suffix++; }
  const output = workbook.addWorksheet(name);
  output.setShowGridlines(false);
  // One category cell for single-bar extent distributions.
  output.getRange("AA1").setValue("Institutions");
  const colours = ["#226481", "#69A8BD", "#E9C46A", "#D47B55", "#CBCFD2"];
  let top = 20;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const categories = source.getRangeByIndexes(block.row, block.column, block.labels.length, 1);
    const percentages = source.getRangeByIndexes(block.row, block.column + 3, block.labels.length, 1);
    let type = ExcelScript.ChartType.barClustered;
    if (block.kind === "pie") type = ExcelScript.ChartType.pie;
    if (block.kind === "stacked") type = ExcelScript.ChartType.barStacked100;
    if (block.kind === "column") type = ExcelScript.ChartType.columnClustered;
    const chart = output.addChart(type, percentages);
    // Explicitly define series so counts and base never become plotted series.
    const initialSeries = chart.getSeries();
    for (const series of initialSeries) series.delete();
    if (block.kind === "stacked") {
      for (let j = 0; j < block.labels.length; j++) {
        const series = chart.addChartSeries(block.labels[j]);
        series.setXAxisValues(output.getRange("AA1"));
        series.setValues(source.getRangeByIndexes(block.row + j, block.column + 3, 1, 1));
        series.getFormat().getFill().setSolidColor(colours[j % colours.length]);
      }
    } else {
      const series = chart.addChartSeries("2025–2026");
      series.setXAxisValues(categories);
      series.setValues(percentages);
      if (block.kind === "pie") {
        series.setVaryByCategories(true);
      } else {
        series.getFormat().getFill().setSolidColor(colours[0]);
      }
      const labels = series.getDataLabels();
      labels.setShowValue(block.kind !== "pie");
      labels.setShowPercentage(block.kind === "pie");
      labels.setShowCategoryName(false);
      labels.setShowSeriesName(false);
      labels.setLinkNumberFormat(false);
      labels.setNumberFormat("0.0%");
      labels.setPosition(block.kind === "pie" ? ExcelScript.ChartDataLabelPosition.bestFit : ExcelScript.ChartDataLabelPosition.outsideEnd);
    }
    chart.setName("Survey_" + (i + 1));
    chart.getTitle().setText(block.title + "\n2025–2026 (n = " + block.base + ")");
    chart.getTitle().setVisible(true);
    chart.getTitle().getFormat().getFont().setSize(14);
    chart.getFormat().getFill().setSolidColor("#FFFFFF");
    const legend = chart.getLegend();
    legend.setVisible(block.kind === "pie" || block.kind === "stacked");
    legend.setPosition(block.kind === "pie" ? ExcelScript.ChartLegendPosition.right : ExcelScript.ChartLegendPosition.bottom);
    if (block.kind !== "pie") {
      const axis = chart.getAxes().getValueAxis();
      axis.setMinimumScale(0);
      axis.setMaximumScale(1);
      axis.setNumberFormat("0%");
      if (block.kind !== "column") chart.getAxes().getCategoryAxis().setReversePlotOrder(true);
    }
    const height = block.labels.length > 9 ? 490 : 360;
    chart.setLeft(20);
    chart.setTop(top);
    chart.setWidth(900);
    chart.setHeight(height);
    top += height + 30;
  }
  output.activate();
  console.log("Created " + blocks.length + " editable charts on '" + name + "', linked to '" + source.getName() + "'. Original sheets were not changed.");
}

function normalize(value: string): string {
  return value.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function findQuestions(values: (string | number | boolean)[][], rowOffset: number, colOffset: number): QuestionBlock[] {
  const result: QuestionBlock[] = [];
  for (let r = 1; r < values.length; r++) {
    for (let c = 0; c + 3 < values[r].length; c++) {
      if (!["category", "categorie"].includes(normalize(String(values[r][c])))) continue;
      if (!normalize(String(values[r][c + 2])).startsWith("base")) continue;
      const title = String(values[r - 1][c]).trim();
      if (!title) continue;
      const match = title.match(/^Q\s*(\d+)/i);
      const question = match ? Number(match[1]) : 0;
      const labels: string[] = [];
      const percentages: number[] = [];
      let base = 0;
      // The three overall initiative/capability tables include an extra union
      // total after the options. It is deliberately not a separate chart bar.
      const limit = question === 9 || question === 11 ? 3 : question === 16 ? 4 : 100;
      // Group follow-ups for Q9/Q11 have 14 categories and must stay complete.
      const first = normalize(String(values[r + 1]?.[c] || ""));
      const isOverall = /internal|interne|employees|employes/.test(first);
      const cap = (question === 9 || question === 11) && !isOverall ? 100 : limit;
      for (let k = r + 1; k < values.length && labels.length < cap; k++) {
        const label = String(values[k][c]).trim();
        const count = values[k][c + 1];
        const n = values[k][c + 2];
        const pct = values[k][c + 3];
        if (!label || n === "" || typeof n !== "number") break;
        if (/^any\s/i.test(label)) break;
        if (n <= 0 || typeof count !== "number" || count < 0 || count > n || typeof pct !== "number" || pct < 0 || pct > 1) {
          throw new Error("Invalid count, base or percentage in " + title + ". Percentages must be numeric Excel percentages.");
        }
        if (Math.abs(pct - count / n) > 0.00051) {
          throw new Error("Count/base does not match the percentage in " + title + ": " + label);
        }
        if (base !== 0 && base !== n) throw new Error("Inconsistent base within " + title);
        base = n;
        labels.push(label);
        percentages.push(pct);
      }
      if (!labels.length) continue;
      const cleaned = labels.map(label => normalize(label));
      const yesNo = cleaned.length === 2 && cleaned.some(label => label === "yes" || label === "oui") && cleaned.some(label => label === "no" || label === "non");
      const rating = [1, 3, 8, 12, 14].includes(question) && labels.length === 5;
      const pie = yesNo || (question === 0 && labels.length <= 5);
      const sum = percentages.reduce((a, b) => a + b, 0);
      if ((pie || rating) && Math.abs(sum - 1) > 0.003) throw new Error("Distribution does not total approximately 100%: " + title);
      let kind = "bar";
      if (pie) kind = "pie";
      else if (rating) kind = "stacked";
      else if ([2, 9, 11, 16].includes(question) && labels.length <= 4) kind = "column";
      result.push({title, question, row: rowOffset + r + 1, column: colOffset + c, labels, percentages, base, kind});
    }
  }
  return result;
}
