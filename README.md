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

Answer rows follow the header. Counts, bases and percentages must be Excel numbers, with percentages stored as proportions (for example, `0.75` formatted as `75%`). Blank rows or rows without a numeric base end a block. Multiple blocks can sit side by side.

The script is tailored to the Q1–Q18 survey layout and currently labels charts **2025–2026**. It uses the existing analysis figures and does not recalculate eligibility or resolve questionnaire routing.

## Chart choices

- Pie: Yes/No questions and institution size/type.
- 100% stacked bars: Q1, Q3, Q8, Q12 and Q14 extent ratings.
- Columns: short overall tables for Q2, Q9, Q11 and Q16.
- Horizontal bars: remaining multi-select tables and longer category lists.

Charts stay linked to source cells in the same workbook. Titles and series names are set when the script runs. Each of Q3's three tables gets its own chart.

## Validation and scope

Table detection and source ranges were checked locally against a 33-table layout, including French headers and shifted cell positions. The script has not yet been tested in the target Excel cloud workbook.

Only the script and these instructions are published here. No survey responses or analysis datasets are included.
