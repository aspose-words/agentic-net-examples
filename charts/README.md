# Charts Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Charts category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Charts
- Slug: charts
- Total examples: 39
- Publish-ready successful examples: 39 / 39
- Source run: 20260619_131835_59df5f
- Existing DOCX / export examples: 5
- Native chart API examples: 28
- Stream / batch / input-bootstrap examples: 3
- Validation examples: 3

## Category rules that shaped these examples

- Do not modify a shape as if it were a chart unless shape.HasChart is true.
- Do not invent chart APIs, enums, or formatting members that are not part of Aspose.Words.
- Do not remove or add series without validating current series availability and compatible data structure.
- Insert charts with DocumentBuilder.InsertChart using the required ChartType.
- Bootstrap local DOCX inputs when a task mentions an existing file, paragraph, chart, stream, or folder.
- Validate series and point counts explicitly for verification scenarios.
- Balance table insertion state before inserting a chart into a cell.
- Initialize all non-nullable reference type properties to avoid CS8618 warnings.
- Avoid CS8600, CS8602, and CS8604 by guarding maybe-null values before dereference or assignment.
- Declare nullable locals when a value may legitimately be null, then null-check before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\charts\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `charts/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\charts\load-an-existing-docx-file-locate-a-chart-shape-by-its-title-and-replace-its-data-source.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-an-existing-docx-file-locate-a-chart-shape-by-its-title-and-replace-its-data-source.cs` | Load an existing DOCX file, locate a chart shape by its title, and replace its data source. | Existing DOCX / export | docx | mcp |
| 2 | `insert-a-column-chart-into-a-new-document-using-documentbuilder-insertchart-with-default-d.cs` | Insert a column chart into a new document using DocumentBuilder.InsertChart with default data. | Native chart API | docx | mcp |
| 3 | `add-a-scatter-chart-to-an-existing-paragraph-by-calling-documentbuilder-insertchart-with-t.cs` | Add a scatter chart to an existing paragraph by calling DocumentBuilder.InsertChart with the appropriate overload. | Existing DOCX / export | docx | mcp |
| 4 | `insert-a-chart-using-a-two-dimensional-array-as-a-custom-data-source-mapping-series-and-ca.cs` | Insert a chart using a two-dimensional array as a custom data source, mapping series and categories. | Native chart API | docx | mcp |
| 5 | `insert-a-chart-into-a-table-cell-and-ensure-it-scales-proportionally-with-the-cell-dimensi.cs` | Insert a chart into a table cell and ensure it scales proportionally with the cell dimensions. | Validation | docx | mcp |
| 6 | `clone-a-chart-shape-from-one-document-section-and-insert-the-duplicate-into-another-paragr.cs` | Clone a chart shape from one document section and insert the duplicate into another paragraph. | Native chart API | docx | mcp |
| 7 | `retrieve-the-shape-chart-object-from-an-inserted-chart-and-modify-its-title-text-programma.cs` | Retrieve the Shape.Chart object from an inserted chart and modify its title text programmatically. | Native chart API | docx | mcp |
| 8 | `create-a-new-chart-series-set-its-values-via-the-series-values-property-and-assign-custom.cs` | Create a new chart series, set its values via the series.Values property, and assign custom category labels. | Native chart API | docx | mcp |
| 9 | `use-chartseriescollection-add-overload-accepting-a-name-and-values-to-create-a-labeled-ser.cs` | Use ChartSeriesCollection.Add overload accepting a name and values to create a labeled series in one step. | Native chart API | docx | mcp |
| 10 | `add-multiple-chartdatapoint-objects-to-a-series-and-set-each-point-s-color-using-the-fill.cs` | Add multiple ChartDataPoint objects to a series and set each point's color using the Fill property. | Stream / batch / input-bootstrap | docx | mcp |
| 11 | `change-the-chart-type-from-column-to-line-after-populating-data-to-demonstrate-dynamic-tra.cs` | Change the chart type from column to line after populating data to demonstrate dynamic transformation. | Native chart API | docx | mcp |
| 12 | `add-a-trendline-to-a-scatter-chart-series-and-configure-its-type-color-and-display-equatio.cs` | Add a trendline to a scatter chart series and configure its type, color, and display equation. | Native chart API | docx | mcp |
| 13 | `configure-chartdatalabel-number-format-to-display-percentages-with-one-decimal-place-for-a.cs` | Configure ChartDataLabel number format to display percentages with one decimal place for all series data points. | Native chart API | docx | mcp |
| 14 | `align-multi-line-chart-data-labels-to-the-center-and-enable-text-wrapping-for-better-reada.cs` | Align multi-line chart data labels to the center and enable text wrapping for better readability. | Native chart API | docx | mcp |
| 15 | `define-default-chartdatalabel-options-to-apply-consistent-font-size-and-color-across-all-c.cs` | Define default ChartDataLabel options to apply consistent font size and color across all chart series. | Native chart API | docx | mcp |
| 16 | `customize-the-chart-s-data-label-font-to-use-a-specific-typeface-size-and-bold-styling-for.cs` | Customize the chart's data label font to use a specific typeface, size, and bold styling for emphasis. | Native chart API | docx | mcp |
| 17 | `enable-data-label-leader-lines-for-a-pie-chart-and-customize-their-length-for-better-place.cs` | Enable data label leader lines for a pie chart and customize their length for better placement. | Native chart API | docx | mcp |
| 18 | `configure-the-chart-to-display-data-labels-only-for-points-exceeding-a-specified-threshold.cs` | Configure the chart to display data labels only for points exceeding a specified threshold value. | Native chart API | docx | mcp |
| 19 | `set-major-gridlines-visibility-on-the-primary-x-axis-and-customize-their-line-color-and-th.cs` | Set major gridlines visibility on the primary X-axis and customize their line color and thickness. | Native chart API | docx | mcp |
| 20 | `adjust-the-primary-y-axis-scaling-to-fixed-minimum-and-maximum-values-and-set-major-unit-i.cs` | Adjust the primary Y-axis scaling to fixed minimum and maximum values and set major unit intervals. | Native chart API | docx | mcp |
| 21 | `set-display-units-for-the-secondary-x-axis-to-thousands-and-format-axis-labels-with-a-cust.cs` | Set display units for the secondary X-axis to thousands and format axis labels with a custom number format. | Native chart API | docx | mcp |
| 22 | `set-the-secondary-y-axis-number-format-to-currency-with-two-decimal-places-for-financial-c.cs` | Set the secondary Y-axis number format to currency with two decimal places for financial charts. | Native chart API | docx | mcp |
| 23 | `apply-a-solid-fill-color-to-the-chart-plot-area-and-add-a-gradient-overlay-for-visual-dept.cs` | Apply a solid fill color to the chart plot area and add a gradient overlay for visual depth. | Native chart API | docx | mcp |
| 24 | `programmatically-set-the-chart-s-background-fill-to-a-semi-transparent-color-to-create-a-w.cs` | Programmatically set the chart's background fill to a semi-transparent color to create a watermark effect. | Native chart API | docx | mcp |
| 25 | `add-a-border-stroke-to-the-chart-legend-with-specified-thickness-and-dash-style-for-emphas.cs` | Add a border stroke to the chart legend with specified thickness and dash style for emphasis. | Native chart API | docx | mcp |
| 26 | `adjust-the-chart-legend-position-to-the-top-right-corner-and-set-its-background-fill-to-li.cs` | Adjust the chart legend position to the top right corner and set its background fill to light gray. | Native chart API | docx | mcp |
| 27 | `update-the-chart-title-text-and-toggle-legend-visibility-based-on-user-preferences.cs` | Update the chart title text and toggle legend visibility based on user preferences. | Native chart API | docx | mcp |
| 28 | `apply-a-predefined-chart-style-template-to-a-newly-inserted-chart-to-ensure-consistent-vis.cs` | Apply a predefined chart style template to a newly inserted chart to ensure consistent visual branding. | Validation | docx | mcp |
| 29 | `apply-a-three-dimensional-rotation-effect-to-a-column-chart-to-enhance-visual-perspective.cs` | Apply a three-dimensional rotation effect to a column chart to enhance visual perspective. | Native chart API | docx | mcp |
| 30 | `programmatically-change-the-chart-s-plot-area-border-to-a-dashed-line-with-specific-color.cs` | Programmatically change the chart's plot area border to a dashed line with specific color and width. | Native chart API | docx | mcp |
| 31 | `programmatically-hide-the-chart-s-plot-area-border-while-keeping-axis-lines-visible-for-a.cs` | Programmatically hide the chart's plot area border while keeping axis lines visible for a clean appearance. | Native chart API | docx | mcp |
| 32 | `enable-automatic-resizing-of-chart-elements-when-the-document-page-size-changes-to-maintai.cs` | Enable automatic resizing of chart elements when the document page size changes to maintain layout integrity. | Native chart API | docx | mcp |
| 33 | `retrieve-existing-chart-series-modify-their-data-points-and-refresh-the-chart-display.cs` | Retrieve existing chart series, modify their data points, and refresh the chart display. | Existing DOCX / export | docx | mcp |
| 34 | `remove-a-specific-series-from-a-chart-using-chartseriescollection-removeat-with-the-correc.cs` | Remove a specific series from a chart using ChartSeriesCollection.RemoveAt with the correct index. | Native chart API | docx | mcp |
| 35 | `validate-that-a-chart-contains-the-expected-number-of-series-and-data-points-before-saving.cs` | Validate that a chart contains the expected number of series and data points before saving. | Validation | docx | mcp |
| 36 | `validate-that-all-chart-series-have-matching-category-counts-to-prevent-data-misalignment.cs` | Validate that all chart series have matching category counts to prevent data misalignment errors during rendering. | Existing DOCX / export | docx | mcp |
| 37 | `export-a-word-document-containing-multiple-charts-to-pdf-while-preserving-chart-formatting.cs` | Export a Word document containing multiple charts to PDF while preserving chart formatting and data labels. | Stream / batch / input-bootstrap | pdf, docx | mcp |
| 38 | `implement-error-handling-to-catch-exceptions-when-inserting-a-chart-into-a-read-only-docum.cs` | Implement error handling to catch exceptions when inserting a chart into a read-only document stream. | Existing DOCX / export | docx | mcp |
| 39 | `batch-process-a-folder-of-word-files-adding-a-predefined-bar-chart-to-each-document-s-firs.cs` | Batch process a folder of Word files, adding a predefined bar chart to each document's first page. | Stream / batch / input-bootstrap | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Invalid chart shape assumption

- Symptom: Code treats a generic Shape as a chart without checking HasChart.
- Fix: Validate shape.HasChart before accessing the Chart property.

### Series/data misalignment

- Symptom: Category count and series values do not match expected chart structure.
- Fix: Clear default data if needed and add aligned categories and series values deterministically.

### Invalid table insertion state

- Symptom: Chart insertion in a table cell fails because the builder is left in an unbalanced table state.
- Fix: Balance StartTable, InsertCell, EndRow, and EndTable before saving.

### Missing local bootstrap inputs

- Symptom: Existing DOCX, folder, or stream scenarios assume inputs already exist.
- Fix: Create local sample input documents or folders inside the example first.

### Nullable warnings

- Symptom: CS8600, CS8602, or CS8604 around maybe-null chart shapes or located nodes.
- Fix: Use nullable locals and guard maybe-null values before dereference.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
