---
name: charts
description: Verified C# examples for charts scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Charts

## Purpose

Provide an AI coding assistant with concise guidance to generate correct, compilable C# snippets that work with Aspose.Words 26.11.0 chart features, ensuring proper document bootstrap, API usage, and null-safety.

## Non-negotiable conventions

- Never treat a generic Shape as a chart without first checking Shape.HasChart.
- Do not reference ChartType, Chart, or any chart members that are not present in the installed Aspose.Words version.
- Do not add or remove series without first clearing the default series collection.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Existing DOCX and export workflow**: 5 examples
- **Native chart API workflow**: 28 examples
- **Validation workflow**: 3 examples
- **Stream and batch I/O workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Ensure the code compiles against Aspose.Words 26.11.0 and that every accessed chart member exists in that version.

## File-to-task reference

- `load-an-existing-docx-file-locate-a-chart-shape-by-its-title-and-replace-its-data-source.cs`
  - Task: Load an existing DOCX file, locate a chart shape by its title, and replace its data source.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-column-chart-into-a-new-document-using-documentbuilder-insertchart-with-default-d.cs`
  - Task: Insert a column chart into a new document using DocumentBuilder.InsertChart with default data.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `add-a-scatter-chart-to-an-existing-paragraph-by-calling-documentbuilder-insertchart-with-t.cs`
  - Task: Add a scatter chart to an existing paragraph by calling DocumentBuilder.InsertChart with the appropriate overload.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-chart-using-a-two-dimensional-array-as-a-custom-data-source-mapping-series-and-ca.cs`
  - Task: Insert a chart using a two-dimensional array as a custom data source, mapping series and categories.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-chart-into-a-table-cell-and-ensure-it-scales-proportionally-with-the-cell-dimensi.cs`
  - Task: Insert a chart into a table cell and ensure it scales proportionally with the cell dimensions.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: mcp
- `clone-a-chart-shape-from-one-document-section-and-insert-the-duplicate-into-another-paragr.cs`
  - Task: Clone a chart shape from one document section and insert the duplicate into another paragraph.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-shape-chart-object-from-an-inserted-chart-and-modify-its-title-text-programma.cs`
  - Task: Retrieve the Shape.Chart object from an inserted chart and modify its title text programmatically.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `create-a-new-chart-series-set-its-values-via-the-series-values-property-and-assign-custom.cs`
  - Task: Create a new chart series, set its values via the series.Values property, and assign custom category labels.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `use-chartseriescollection-add-overload-accepting-a-name-and-values-to-create-a-labeled-ser.cs`
  - Task: Use ChartSeriesCollection.Add overload accepting a name and values to create a labeled series in one step.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `add-multiple-chartdatapoint-objects-to-a-series-and-set-each-point-s-color-using-the-fill.cs`
  - Task: Add multiple ChartDataPoint objects to a series and set each point's color using the Fill property.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: llm
- `change-the-chart-type-from-column-to-line-after-populating-data-to-demonstrate-dynamic-tra.cs`
  - Task: Change the chart type from column to line after populating data to demonstrate dynamic transformation.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `add-a-trendline-to-a-scatter-chart-series-and-configure-its-type-color-and-display-equatio.cs`
  - Task: Add a trendline to a scatter chart series and configure its type, color, and display equation.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `configure-chartdatalabel-number-format-to-display-percentages-with-one-decimal-place-for-a.cs`
  - Task: Configure ChartDataLabel number format to display percentages with one decimal place for all series data points.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `align-multi-line-chart-data-labels-to-the-center-and-enable-text-wrapping-for-better-reada.cs`
  - Task: Align multi-line chart data labels to the center and enable text wrapping for better readability.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `define-default-chartdatalabel-options-to-apply-consistent-font-size-and-color-across-all-c.cs`
  - Task: Define default ChartDataLabel options to apply consistent font size and color across all chart series.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `customize-the-chart-s-data-label-font-to-use-a-specific-typeface-size-and-bold-styling-for.cs`
  - Task: Customize the chart's data label font to use a specific typeface, size, and bold styling for emphasis.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `enable-data-label-leader-lines-for-a-pie-chart-and-customize-their-length-for-better-place.cs`
  - Task: Enable data label leader lines for a pie chart and customize their length for better placement.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `configure-the-chart-to-display-data-labels-only-for-points-exceeding-a-specified-threshold.cs`
  - Task: Configure the chart to display data labels only for points exceeding a specified threshold value.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `set-major-gridlines-visibility-on-the-primary-x-axis-and-customize-their-line-color-and-th.cs`
  - Task: Set major gridlines visibility on the primary X-axis and customize their line color and thickness.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `adjust-the-primary-y-axis-scaling-to-fixed-minimum-and-maximum-values-and-set-major-unit-i.cs`
  - Task: Adjust the primary Y-axis scaling to fixed minimum and maximum values and set major unit intervals.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `set-display-units-for-the-secondary-x-axis-to-thousands-and-format-axis-labels-with-a-cust.cs`
  - Task: Set display units for the secondary X-axis to thousands and format axis labels with a custom number format.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `set-the-secondary-y-axis-number-format-to-currency-with-two-decimal-places-for-financial-c.cs`
  - Task: Set the secondary Y-axis number format to currency with two decimal places for financial charts.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: llm
- `apply-a-solid-fill-color-to-the-chart-plot-area-and-add-a-gradient-overlay-for-visual-dept.cs`
  - Task: Apply a solid fill color to the chart plot area and add a gradient overlay for visual depth.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `programmatically-set-the-chart-s-background-fill-to-a-semi-transparent-color-to-create-a-w.cs`
  - Task: Programmatically set the chart's background fill to a semi-transparent color to create a watermark effect.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `add-a-border-stroke-to-the-chart-legend-with-specified-thickness-and-dash-style-for-emphas.cs`
  - Task: Add a border stroke to the chart legend with specified thickness and dash style for emphasis.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `adjust-the-chart-legend-position-to-the-top-right-corner-and-set-its-background-fill-to-li.cs`
  - Task: Adjust the chart legend position to the top right corner and set its background fill to light gray.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `update-the-chart-title-text-and-toggle-legend-visibility-based-on-user-preferences.cs`
  - Task: Update the chart title text and toggle legend visibility based on user preferences.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: llm
- `apply-a-predefined-chart-style-template-to-a-newly-inserted-chart-to-ensure-consistent-vis.cs`
  - Task: Apply a predefined chart style template to a newly inserted chart to ensure consistent visual branding.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: existing_repo
- `apply-a-three-dimensional-rotation-effect-to-a-column-chart-to-enhance-visual-perspective.cs`
  - Task: Apply a three-dimensional rotation effect to a column chart to enhance visual perspective.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: llm
- `programmatically-change-the-chart-s-plot-area-border-to-a-dashed-line-with-specific-color.cs`
  - Task: Programmatically change the chart's plot area border to a dashed line with specific color and width.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `programmatically-hide-the-chart-s-plot-area-border-while-keeping-axis-lines-visible-for-a.cs`
  - Task: Programmatically hide the chart's plot area border while keeping axis lines visible for a clean appearance.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: existing_repo
- `enable-automatic-resizing-of-chart-elements-when-the-document-page-size-changes-to-maintai.cs`
  - Task: Enable automatic resizing of chart elements when the document page size changes to maintain layout integrity.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-existing-chart-series-modify-their-data-points-and-refresh-the-chart-display.cs`
  - Task: Retrieve existing chart series, modify their data points, and refresh the chart display.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: existing_repo
- `remove-a-specific-series-from-a-chart-using-chartseriescollection-removeat-with-the-correc.cs`
  - Task: Remove a specific series from a chart using ChartSeriesCollection.RemoveAt with the correct index.
  - Workflow: native-chart-api
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-a-chart-contains-the-expected-number-of-series-and-data-points-before-saving.cs`
  - Task: Validate that a chart contains the expected number of series and data points before saving.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-all-chart-series-have-matching-category-counts-to-prevent-data-misalignment.cs`
  - Task: Validate that all chart series have matching category counts to prevent data misalignment errors during rendering.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: llm
- `export-a-word-document-containing-multiple-charts-to-pdf-while-preserving-chart-formatting.cs`
  - Task: Export a Word document containing multiple charts to PDF while preserving chart formatting and data labels.
  - Workflow: stream-batch-io
  - Outputs: pdf, docx
  - Selected engine: existing_repo
- `implement-error-handling-to-catch-exceptions-when-inserting-a-chart-into-a-read-only-docum.cs`
  - Task: Implement error handling to catch exceptions when inserting a chart into a read-only document stream.
  - Workflow: existing-docx-and-export
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-a-folder-of-word-files-adding-a-predefined-bar-chart-to-each-document-s-firs.cs`
  - Task: Batch process a folder of Word files, adding a predefined bar chart to each document's first page.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Invalid chart type reference**
  - Symptom: The compiler reports CS0103 for 'ChartType' and CS0246 for missing 'Chart' namespace.
  - Root cause: The example uses a ChartType enum or Chart class that is not part of the referenced Aspose.Words version.
  - Preferred fix: Import Aspose.Words.Drawing.Charts and use only the ChartType values documented for version 26.11.0; replace missing types with the correct enum members.
  - Prevention: Consult the Aspose.Words API reference for the exact enum names before writing code and add the required using directive.

- **Missing chart API members**
  - Symptom: CS1061 errors appear for members such as CategoryAxis, DataLabel, Trendlines, Fill, or NumberFormat.
  - Root cause: The code attempts to access properties or methods that were removed or never existed in the current library.
  - Preferred fix: Remove calls to unavailable members and use the supported alternatives, such as ChartSeries.HasDataLabels, ChartSeries.DataLabels, and ChartSeries.DataPoints[i].Format.Fill for coloring.
  - Prevention: Reference the official Aspose.Words chart documentation and verify member availability with IntelliSense or the API browser.

- **Assuming shape contains a chart**
  - Symptom: Runtime or compile-time errors occur when accessing shape.Chart without checking HasChart.
  - Root cause: The example treats any Shape as a chart, violating the strict API ban.
  - Preferred fix: Guard chart access with if (shape.HasChart) { Chart chart = shape.Chart; ... } and throw a clear exception when the condition fails.
  - Prevention: Always include a HasChart check before dereferencing the Chart property.

- **Incorrect series/data API usage**
  - Symptom: CS1061 errors for missing Values, CategoryNames, or Add overloads; data misalignment warnings during rendering.
  - Root cause: The code uses outdated overload signatures or mismatched arrays for categories and values.
  - Preferred fix: Clear the default series collection, then add series using the overload that accepts (string seriesName, string[] categories, double[] values) or (string seriesName, double[] xValues, double[] yValues) for scatter charts.
  - Prevention: Validate that the length of category and value arrays match before calling Add, and always clear existing series first.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\charts\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically verify that all chart members used in examples still exist in the latest Aspose.Words release and update the code accordingly.
