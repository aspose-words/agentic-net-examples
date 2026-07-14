using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartSeriesExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a newly inserted chart.
        chart.Series.Clear();

        // Define custom category labels and corresponding values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 1500, 2300, 1800, 2100 };

        // Add a new series with the custom categories and values.
        chart.Series.Add("Fiscal Year 2023", categories, values);

        // Save the document to the local file system.
        doc.Save("ChartSeriesExample.docx");
    }
}
