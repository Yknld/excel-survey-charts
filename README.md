# Excel survey charts

An Office Script that creates editable Excel charts from survey analysis tables in English or French.

## Run in Excel

1. Open `create_survey_charts.ts` in this repository and copy the entire file (use GitHub's copy button or **Raw**).
2. Open your workbook and select **Analysis** or **Analysis FR**.
3. Choose **Automate → New Script**.
4. Replace the starter code with the copied script, save it, and click **Run**.

The script creates a new **Charts EN Auto** or **Charts FR Auto** sheet. Running it again creates another numbered sheet. It does not overwrite existing sheets or charts.

## Required table layout

Each question block must have its title immediately above a four-column header row:

| Category / Catégorie | Institutions | Base (n) | 2025–2026 |
|---|---:|---:|---:|

Answer rows follow the header. Counts and bases can be numbers or text numbers. Percentages can be numeric proportions (for example, `0.75` formatted as `75%`) or pasted text such as `75%`, `43.5%`, `43,5 %`, or `0,435`. Common non-breaking spaces are supported. Blanks are not treated as zero. Blank rows or rows with a blank base end a block. Multiple blocks can sit side by side.

The script checks percentages against counts and bases before creating charts. It creates formula-linked numeric helper cells in columns AA–AB of the new chart sheet, calculating exact proportions from counts and bases. Keep those helper cells: they let charts work with text-formatted source numbers and update when source counts change. The source sheet is not modified. The helper formulas do not repeat the initial validation after later edits; rerun the script to validate changed data.

The script is tailored to the Q1–Q18 survey layout and currently labels charts **2025–2026**. It uses the existing analysis figures and does not recalculate eligibility or resolve questionnaire routing.

## Chart choices

- Pie: Yes/No questions and institution size/type.
- 100% stacked bars: Q1, Q3, Q8, Q12 and Q14 extent ratings.
- Columns: short overall tables for Q2, Q9, Q11 and Q16.
- Horizontal bars: remaining multi-select tables and longer category lists.

Charts stay linked to source cells in the same workbook. Titles and series names are set when the script runs. Each of Q3's three tables gets its own chart.

## Validation and scope

The script passes a TypeScript check against Microsoft's Office Scripts release API declarations. Local tests cover number parsing, French decimal commas and spaces, rejection of invalid data before sheet creation, all four chart types, source links, repeat runs, source preservation, and all 33 tables in the original layout. Workbook calls are tested with a simulated workbook; this is not a live Excel runtime test. The corrected script still needs a run in the target Excel cloud workbook.

### Run the local tests

Install Node.js and the development dependency, then download Microsoft's declarations and run:

```sh
npm install
curl -fL https://raw.githubusercontent.com/OfficeDev/office-scripts-docs-reference/main/generate-docs/api-extractor-inputs-excelscript-release/excelscript.d.ts -o excel.d.ts
npm test
```

The tests use synthetic data. An optional local analysis JSON can be supplied to `node test.cjs /path/to/analysis_tables_data.json`; no real response data is included in this repository.

Individual survey responses are not published here.

## Q19–Q20 dataset history

The Q19–Q20 theme-statistics dataset has been removed from the current version. Its previous version remains available in [Git history](https://github.com/Yknld/excel-survey-charts/commits/main/analysis/2025-2026/Q19_Q20_theme_statistics.csv).
