using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;

class ChartSeriesValidator
{
    // Validates that every series in the chart has the same number of categories (X values)
    // and that each series has an equal number of Y values.
    private static void ValidateSeriesCategoryCounts(Chart chart)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));

        int? expectedCategoryCount = null;

        foreach (ChartSeries series in chart.Series)
        {
            // Number of categories for the current series.
            int categoryCount = series.XValues.Count;

            // Ensure Y values count matches X values count for the series itself.
            if (series.YValues.Count != categoryCount)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' has mismatched X/Y counts (X: {categoryCount}, Y: {series.YValues.Count}).");

            // Establish the expected category count from the first series.
            if (expectedCategoryCount == null)
            {
                expectedCategoryCount = categoryCount;
            }
            else if (categoryCount != expectedCategoryCount.Value)
            {
                throw new InvalidOperationException(
                    $"Series '{series.Name}' has {categoryCount} categories, expected {expectedCategoryCount.Value}.");
            }
        }
    }

    static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear any demo data.
        chart.Series.Clear();

        // Define categories (must be the same for all series).
        string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };

        // Add series with matching category counts.
        chart.Series.Add("Series 1", categories, new[] { 10.0, 20.0, 30.0, 40.0 });
        chart.Series.Add("Series 2", categories, new[] { 15.0, 25.0, 35.0, 45.0 });
        chart.Series.Add("Series 3", categories, new[] { 12.0, 22.0, 32.0, 42.0 });

        // Perform validation – will throw if any series is inconsistent.
        ValidateSeriesCategoryCounts(chart);

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "ValidatedChart.docx"));
    }
}
