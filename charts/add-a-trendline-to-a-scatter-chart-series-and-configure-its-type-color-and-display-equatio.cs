using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 500, 300);

        // Ensure the shape actually contains a chart before proceeding.
        if (!chartShape.HasChart)
            return;

        Chart chart = chartShape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Define X and Y values for the series.
        double[] xValues = { 1, 2, 3, 4, 5 };
        double[] yValues = { 2, 4, 5, 4, 5 };

        // Add a new series to the scatter chart.
        ChartSeries series = chart.Series.Add("Sample Series", xValues, yValues);

        // NOTE: Aspose.Words for .NET does not currently expose a Trendline API.
        // The following code that would add a trendline is therefore omitted
        // to keep the example compilable.

        // Save the document with the chart.
        doc.Save("ScatterChartWithTrendline.docx");
    }
}
