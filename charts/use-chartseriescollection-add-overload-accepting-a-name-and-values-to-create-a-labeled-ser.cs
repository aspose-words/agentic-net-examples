using System;
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

        // Insert a histogram chart. Width and height are in points.
        Shape chartShape = builder.InsertChart(ChartType.Histogram, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a new chart.
        chart.Series.Clear();

        // Add a new series with a name and a set of values in one step.
        // This uses the ChartSeriesCollection.Add(string, double[]) overload.
        chart.Series.Add("Sample Series", new double[] { 10, 20, 30, 40, 25 });

        // Save the document to the working directory.
        doc.Save("ChartSeriesAdd.docx");
    }
}
