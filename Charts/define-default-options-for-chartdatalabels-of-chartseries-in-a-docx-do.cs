using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add a sample series with three categories.
        chart.Series.Add(
            "Series 1",
            new[] { "Category 1", "Category 2", "Category 3" },
            new[] { 4.0, 5.0, 6.0 });

        // Get the first (and only) series.
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Configure default options for all data labels in the series.
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowValue = true;                     // Show the numeric value.
        dataLabels.ShowCategoryName = true;              // Show the category name.
        dataLabels.ShowSeriesName = false;               // Do not show the series name.
        dataLabels.ShowPercentage = false;               // Not a pie chart – keep false.
        dataLabels.ShowLegendKey = false;                // No legend key.
        dataLabels.ShowLeaderLines = false;              // No leader lines.
        dataLabels.Separator = ", ";                     // Separator between parts.
        dataLabels.Font.Color = Color.White;             // Font color for the whole series.
        dataLabels.Position = ChartDataLabelPosition.InsideBase; // Default position.
        dataLabels.Orientation = ShapeTextOrientation.VerticalFarEast; // Text orientation.
        dataLabels.Rotation = -45;                       // Rotate the label.

        // Example of overriding a single label (optional).
        dataLabels[0].Font.Color = Color.DarkRed;
        dataLabels[0].Position = ChartDataLabelPosition.OutsideEnd;

        // Save the document.
        doc.Save("ChartDataLabelsDefaults.docx");
    }
}
