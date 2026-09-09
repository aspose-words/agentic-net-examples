using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // Needed for the Shape class
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds.
        chart.Series.Clear();

        // Define categories (X‑axis labels) and corresponding Y values.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        double[] values = { 10, 20, 30 };

        // Add a new series with the categories and values.
        ChartSeries series = chart.Series.Add("Series 1", categories, values);

        // Set a distinct fill color for each data point in the series.
        series.DataPoints[0].Format.Fill.Color = Color.Red;    // First point – red
        series.DataPoints[1].Format.Fill.Color = Color.Green; // Second point – green
        series.DataPoints[2].Format.Fill.Color = Color.Blue;  // Third point – blue

        // Save the document containing the customized chart.
        doc.Save("AddDataPointsColors.docx");
    }
}
