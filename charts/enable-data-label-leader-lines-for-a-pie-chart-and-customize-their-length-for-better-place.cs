using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        ChartSeries series = chart.Series.Add(
            "Sales",
            new[] { "Product A", "Product B", "Product C" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;

        // Show leader lines and values in the data labels.
        dataLabels.ShowLeaderLines = true;
        dataLabels.ShowValue = true;

        // Position the labels outside the pie slices – this effectively lengthens the leader lines.
        dataLabels.Position = ChartDataLabelPosition.OutsideEnd;

        // Optionally customize the separator between displayed parts.
        dataLabels.Separator = "; ";

        // Save the document to the local file system.
        doc.Save("PieChartWithLeaderLines.docx");
    }
}
