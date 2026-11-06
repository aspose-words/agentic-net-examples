# Charts Examples for Aspose.Words for .NET

The Charts category of Aspose.Words supplies native APIs for creating, customizing, and exporting chart objects inside Word documents. The verified examples illustrate inserting various chart types, populating data from arrays, cloning chart shapes, embedding charts in table cells, and dynamically swapping chart types while respecting the library's supported API surface.

## Snapshot

- Category: Charts
- Slug: charts
- Total examples: 39
- Publish-ready successful examples: 39 / 39
- Source run: 20261102_031946_08fb5b
- Existing DOCX and export examples: 5
- Native chart API examples: 28
- Validation examples: 3
- Stream and batch I/O examples: 3

## Category rules that shaped these examples

- Do not treat a generic Shape as a chart unless shape.HasChart is true.
- Do not reference members that are not part of Aspose.Words.Chart, such as CategoryAxis, ChartLegend.IsVisible, or Trendline APIs.
- Do not add or remove series without first clearing or validating the existing series collection.
- Insert charts with DocumentBuilder.InsertChart using the required ChartType enumeration.
- When a task mentions an existing file, paragraph, chart, stream, or folder, create a local sample DOCX or directory at the beginning of the example.
- Validate that each series has a matching number of category labels and values before saving.
- Balance table state (StartTable, InsertCell, EndRow, EndTable) before document save.
- Initialize all non-nullable reference-type fields or use the null-forgiving operator to satisfy the compiler.
- Guard any potentially null nodes (e.g., builder.CurrentParagraph?.ParentNode) with explicit null checks before dereferencing.
- Declare locals as nullable when a value may be absent and perform a null-check prior to use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\charts\load-an-existing-docx-file-locate-a-chart-shape-by-its-title-and-replace-its-data-source.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-an-existing-docx-file-locate-a-chart-shape-by-its-title-and-replace-its-data-source.cs` | Load an existing DOCX file, locate a chart shape by its title, and replace its data source. | existing-docx-and-export | docx | mcp |
| 2 | `insert-a-column-chart-into-a-new-document-using-documentbuilder-insertchart-with-default-d.cs` | Insert a column chart into a new document using DocumentBuilder.InsertChart with default data. | native-chart-api | docx | mcp |
| 3 | `add-a-scatter-chart-to-an-existing-paragraph-by-calling-documentbuilder-insertchart-with-t.cs` | Add a scatter chart to an existing paragraph by calling DocumentBuilder.InsertChart with the appropriate overload. | existing-docx-and-export | docx | mcp |
| 4 | `insert-a-chart-using-a-two-dimensional-array-as-a-custom-data-source-mapping-series-and-ca.cs` | Insert a chart using a two-dimensional array as a custom data source, mapping series and categories. | native-chart-api | docx | llm |
| 5 | `insert-a-chart-into-a-table-cell-and-ensure-it-scales-proportionally-with-the-cell-dimensi.cs` | Insert a chart into a table cell and ensure it scales proportionally with the cell dimensions. | validation | docx | mcp |
| 6 | `clone-a-chart-shape-from-one-document-section-and-insert-the-duplicate-into-another-paragr.cs` | Clone a chart shape from one document section and insert the duplicate into another paragraph. | native-chart-api | docx | mcp |
| 7 | `retrieve-the-shape-chart-object-from-an-inserted-chart-and-modify-its-title-text-programma.cs` | Retrieve the Shape.Chart object from an inserted chart and modify its title text programmatically. | native-chart-api | docx | mcp |
| 8 | `create-a-new-chart-series-set-its-values-via-the-series-values-property-and-assign-custom.cs` | Create a new chart series, set its values via the series.Values property, and assign custom category labels. | native-chart-api | docx | existing_repo |
| 9 | `use-chartseriescollection-add-overload-accepting-a-name-and-values-to-create-a-labeled-ser.cs` | Use ChartSeriesCollection.Add overload accepting a name and values to create a labeled series in one step. | native-chart-api | docx | mcp |
| 10 | `add-multiple-chartdatapoint-objects-to-a-series-and-set-each-point-s-color-using-the-fill.cs` | Add multiple ChartDataPoint objects to a series and set each point's color using the Fill property. | stream-batch-io | docx | mcp |
| 11 | `change-the-chart-type-from-column-to-line-after-populating-data-to-demonstrate-dynamic-tra.cs` | Change the chart type from column to line after populating data to demonstrate dynamic transformation. | native-chart-api | docx | existing_repo |
| 12 | `add-a-trendline-to-a-scatter-chart-series-and-configure-its-type-color-and-display-equatio.cs` | Add a trendline to a scatter chart series and configure its type, color, and display equation. | native-chart-api | docx | existing_repo |
| 13 | `configure-chartdatalabel-number-format-to-display-percentages-with-one-decimal-place-for-a.cs` | Configure ChartDataLabel number format to display percentages with one decimal place for all series data points. | native-chart-api | docx | existing_repo |
| 14 | `align-multi-line-chart-data-labels-to-the-center-and-enable-text-wrapping-for-better-reada.cs` | Align multi-line chart data labels to the center and enable text wrapping for better readability. | native-chart-api | docx | existing_repo |
| 15 | `define-default-chartdatalabel-options-to-apply-consistent-font-size-and-color-across-all-c.cs` | Define default ChartDataLabel options to apply consistent font size and color across all chart series. | native-chart-api | docx | existing_repo |
| 16 | `customize-the-chart-s-data-label-font-to-use-a-specific-typeface-size-and-bold-styling-for.cs` | Customize the chart's data label font to use a specific typeface, size, and bold styling for emphasis. | native-chart-api | docx | existing_repo |
| 17 | `enable-data-label-leader-lines-for-a-pie-chart-and-customize-their-length-for-better-place.cs` | Enable data label leader lines for a pie chart and customize their length for better placement. | native-chart-api | docx | existing_repo |
| 18 | `configure-the-chart-to-display-data-labels-only-for-points-exceeding-a-specified-threshold.cs` | Configure the chart to display data labels only for points exceeding a specified threshold value. | native-chart-api | docx | existing_repo |
| 19 | `set-major-gridlines-visibility-on-the-primary-x-axis-and-customize-their-line-color-and-th.cs` | Set major gridlines visibility on the primary X-axis and customize their line color and thickness. | native-chart-api | docx | llm |
| 20 | `adjust-the-primary-y-axis-scaling-to-fixed-minimum-and-maximum-values-and-set-major-unit-i.cs` | Adjust the primary Y-axis scaling to fixed minimum and maximum values and set major unit intervals. | native-chart-api | docx | existing_repo |
| 21 | `set-display-units-for-the-secondary-x-axis-to-thousands-and-format-axis-labels-with-a-cust.cs` | Set display units for the secondary X-axis to thousands and format axis labels with a custom number format. | native-chart-api | docx | existing_repo |
| 22 | `set-the-secondary-y-axis-number-format-to-currency-with-two-decimal-places-for-financial-c.cs` | Set the secondary Y-axis number format to currency with two decimal places for financial charts. | native-chart-api | docx | existing_repo |
| 23 | `apply-a-solid-fill-color-to-the-chart-plot-area-and-add-a-gradient-overlay-for-visual-dept.cs` | Apply a solid fill color to the chart plot area and add a gradient overlay for visual depth. | native-chart-api | docx | existing_repo |
| 24 | `programmatically-set-the-chart-s-background-fill-to-a-semi-transparent-color-to-create-a-w.cs` | Programmatically set the chart's background fill to a semi-transparent color to create a watermark effect. | native-chart-api | docx | existing_repo |
| 25 | `add-a-border-stroke-to-the-chart-legend-with-specified-thickness-and-dash-style-for-emphas.cs` | Add a border stroke to the chart legend with specified thickness and dash style for emphasis. | native-chart-api | docx | existing_repo |
| 26 | `adjust-the-chart-legend-position-to-the-top-right-corner-and-set-its-background-fill-to-li.cs` | Adjust the chart legend position to the top right corner and set its background fill to light gray. | native-chart-api | docx | llm |
| 27 | `update-the-chart-title-text-and-toggle-legend-visibility-based-on-user-preferences.cs` | Update the chart title text and toggle legend visibility based on user preferences. | native-chart-api | docx | mcp |
| 28 | `apply-a-predefined-chart-style-template-to-a-newly-inserted-chart-to-ensure-consistent-vis.cs` | Apply a predefined chart style template to a newly inserted chart to ensure consistent visual branding. | validation | docx | existing_repo |
| 29 | `apply-a-three-dimensional-rotation-effect-to-a-column-chart-to-enhance-visual-perspective.cs` | Apply a three-dimensional rotation effect to a column chart to enhance visual perspective. | native-chart-api | docx | existing_repo |
| 30 | `programmatically-change-the-chart-s-plot-area-border-to-a-dashed-line-with-specific-color.cs` | Programmatically change the chart's plot area border to a dashed line with specific color and width. | native-chart-api | docx | existing_repo |
| 31 | `programmatically-hide-the-chart-s-plot-area-border-while-keeping-axis-lines-visible-for-a.cs` | Programmatically hide the chart's plot area border while keeping axis lines visible for a clean appearance. | native-chart-api | docx | existing_repo |
| 32 | `enable-automatic-resizing-of-chart-elements-when-the-document-page-size-changes-to-maintai.cs` | Enable automatic resizing of chart elements when the document page size changes to maintain layout integrity. | native-chart-api | docx | mcp |
| 33 | `retrieve-existing-chart-series-modify-their-data-points-and-refresh-the-chart-display.cs` | Retrieve existing chart series, modify their data points, and refresh the chart display. | existing-docx-and-export | docx | mcp |
| 34 | `remove-a-specific-series-from-a-chart-using-chartseriescollection-removeat-with-the-correc.cs` | Remove a specific series from a chart using ChartSeriesCollection.RemoveAt with the correct index. | native-chart-api | docx | mcp |
| 35 | `validate-that-a-chart-contains-the-expected-number-of-series-and-data-points-before-saving.cs` | Validate that a chart contains the expected number of series and data points before saving. | validation | docx | llm |
| 36 | `validate-that-all-chart-series-have-matching-category-counts-to-prevent-data-misalignment.cs` | Validate that all chart series have matching category counts to prevent data misalignment errors during rendering. | existing-docx-and-export | docx | llm |
| 37 | `export-a-word-document-containing-multiple-charts-to-pdf-while-preserving-chart-formatting.cs` | Export a Word document containing multiple charts to PDF while preserving chart formatting and data labels. | stream-batch-io | pdf, docx | existing_repo |
| 38 | `implement-error-handling-to-catch-exceptions-when-inserting-a-chart-into-a-read-only-docum.cs` | Implement error handling to catch exceptions when inserting a chart into a read-only document stream. | existing-docx-and-export | docx | mcp |
| 39 | `batch-process-a-folder-of-word-files-adding-a-predefined-bar-chart-to-each-document-s-firs.cs` | Batch process a folder of Word files, adding a predefined bar chart to each document's first page. | stream-batch-io | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Invalid chart shape assumption

- Symptom: Code accesses Shape.Chart without confirming shape.HasChart, leading to CS1061 errors.
- Root cause: The example treats any Shape as a chart regardless of its actual content.
- Preferred fix: Add a shape.HasChart check and throw a clear exception if the shape does not contain a chart.
- Prevention: Always guard chart access with if (shape.HasChart) { ... } else { throw new InvalidOperationException(...); }.

### Use of non-existent chart members

- Symptom: Compilation fails with CS1061 or CS0246 for members such as CategoryAxis, ChartLegend.IsVisible, Trendline, or DataLabel properties that do not exist in Aspose.Words.
- Root cause: The code references API members that are part of other libraries (e.g., Aspose.Slides) or older versions of Aspose.Words.
- Preferred fix: Replace the invalid members with the supported Aspose.Words equivalents, or omit the unsupported functionality.
- Prevention: Consult the Aspose.Words 26.10.0 documentation for the exact Chart API surface before writing code.

### Table insertion state imbalance

- Symptom: Saving the document throws runtime errors because the table construction (StartTable/InsertCell/EndRow/EndTable) is not balanced.
- Root cause: The example leaves an open table row or cell when inserting a chart.
- Preferred fix: Ensure that every StartTable is paired with EndTable and that each InsertCell is followed by EndRow before saving.
- Prevention: Wrap chart insertion inside a clearly delimited table block and call builder.EndRow() and builder.EndTable() after the chart is placed.

### Missing local bootstrap inputs

- Symptom: Examples that reference external DOCX files, folders, or streams fail at runtime because the inputs do not exist.
- Root cause: The code assumes pre-existing files or directories without creating them in the sample.
- Preferred fix: Create a minimal DOCX, directory, or memory stream at the start of the example and use it for subsequent operations.
- Prevention: Always include a self-contained bootstrap section that generates any required input artifacts.

### Incorrect overload usage for builder navigation

- Symptom: CS7036 errors occur when calling DocumentBuilder.MoveToParagraph without providing the required characterIndex argument.
- Root cause: The example uses an overload that no longer exists or omits required parameters.
- Preferred fix: Use the overload MoveToParagraph(paragraphIndex, characterIndex) with both arguments, or move the builder by referencing the target Paragraph node directly.
- Prevention: Refer to the current Aspose.Words API for the correct overload signatures before invoking navigation methods.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Periodically review Aspose.Words release notes to update the list of supported chart members and adjust the strict API bans accordingly.
