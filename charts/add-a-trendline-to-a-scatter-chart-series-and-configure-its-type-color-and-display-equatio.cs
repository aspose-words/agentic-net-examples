using System;
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
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 500, 400);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds.
        chart.Series.Clear();

        // Define X and Y values for the custom series.
        double[] xValues = { 1, 2, 3, 4, 5 };
        double[] yValues = { 2, 4, 5, 4, 5 };

        // Add the custom series to the chart.
        ChartSeries series = chart.Series.Add("Sample Series", xValues, yValues);

        // NOTE: Aspose.Words for .NET does not currently provide a Trendline API.
        // The following code that attempts to add a trendline would not compile because
        // the required types (ChartTrendline, ChartTrendlineType) are not part of the library.
        // If trendline support is added in a future version, you can uncomment and adjust
        // the code accordingly.

        // ChartTrendline trendline = series.Trendlines.Add(ChartTrendlineType.Linear);
        // trendline.Format.Line.ForeColor = Color.Red;
        // trendline.DisplayEquation = true;

        // Save the document containing the chart.
        doc.Save("ScatterChartWithTrendline.docx");
    }
}
