using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Required for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart related types

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a histogram chart. Histogram charts support the Add(name, values) overload.
        Shape chartShape = builder.InsertChart(ChartType.Histogram, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a new chart.
        chart.Series.Clear();

        // Add a new series with a name and a set of values in one step.
        chart.Series.Add("Sample Series", new double[] { 12.5, 18.3, 9.7, 15.0, 22.1 });

        // Save the document to the local file system.
        doc.Save("ChartSeriesAdd.docx");
    }
}
