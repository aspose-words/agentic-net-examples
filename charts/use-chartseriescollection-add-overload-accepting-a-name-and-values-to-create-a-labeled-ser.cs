using System;
using System.IO;
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

        // Insert a Histogram chart (the type that supports the Add(name, values) overload).
        Shape chartShape = builder.InsertChart(ChartType.Histogram, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a new chart.
        chart.Series.Clear();

        // Define the data values for the new series.
        double[] values = new double[] { 10, 20, 15, 30, 25 };

        // Add a labeled series in one step using the overload that takes a name and values.
        chart.Series.Add("Sample Series", values);

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartSeriesAdd.docx");
        doc.Save(outputPath);
    }
}
