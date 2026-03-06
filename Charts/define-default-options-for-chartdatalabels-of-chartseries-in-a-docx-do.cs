using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartDataLabelsDefaults
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with sample categories and values.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Jan", "Feb", "Mar", "Apr" },
            new[] { 10.0, 20.5, 15.2, 30.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Get the collection that controls the data‑label settings for the whole series.
        ChartDataLabelCollection dataLabels = series.DataLabels;

        // Set default visibility options.
        dataLabels.ShowValue = true;               // Show the numeric value.
        dataLabels.ShowCategoryName = true;        // Show the category (X‑axis) name.
        dataLabels.ShowSeriesName = false;         // Hide the series name.
        dataLabels.ShowLeaderLines = false;        // No leader lines.
        dataLabels.ShowLegendKey = false;          // No legend key.
        dataLabels.ShowPercentage = false;         // Not a pie chart, so hide percentage.
        dataLabels.ShowBubbleSize = false;         // Not a bubble chart.
        dataLabels.ShowDataLabelsRange = false;    // Do not display range values.

        // Set formatting defaults.
        dataLabels.Separator = ", ";               // Default separator between parts.
        dataLabels.NumberFormat.FormatCode = "0.00"; // Two‑decimal numeric format.

        // Font defaults for all labels in the series.
        dataLabels.Font.Size = 10;
        dataLabels.Font.Color = Color.Black;

        // Fill and stroke defaults for all labels in the series.
        dataLabels.Format.Fill.Solid(Color.LightYellow);
        dataLabels.Format.Stroke.Color = Color.DarkGray;

        // Save the document.
        doc.Save("ChartDataLabelsDefaults.docx");
    }
}
