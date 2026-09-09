using System;
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

        // Insert a histogram chart. Histogram charts support the Add(string, double[]) overload.
        Shape chartShape = builder.InsertChart(ChartType.Histogram, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds.
        chart.Series.Clear();

        // Add a labeled series in one step: provide the series name and an array of values.
        chart.Series.Add("Sample Series", new double[] { 10, 20, 15, 30, 25 });

        // Save the document containing the chart.
        doc.Save("ChartSeriesAdd.docx");
    }
}
