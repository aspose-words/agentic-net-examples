using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;        // Chart related types

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Define categories (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 1500, 2300, 1800, 2100 };

        // Add a labeled series in one step using the overload that accepts a name, categories and values.
        chart.Series.Add("Revenue", categories, values);

        // Save the document containing the chart.
        doc.Save("ChartSeriesAdd.docx");
    }
}
