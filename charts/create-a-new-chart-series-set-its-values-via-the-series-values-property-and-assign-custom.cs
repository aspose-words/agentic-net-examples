using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // Needed for the Shape class
using Aspose.Words.Drawing.Charts;   // Chart‑related types

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Define custom category labels for the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Define the numeric values for the series.
        double[] values = { 1500, 2300, 1800, 2100 };

        // Add a new series using the categories and values.
        chart.Series.Add("Fiscal Year 2023", categories, values);

        // Save the document to the working directory.
        doc.Save("ChartSeriesValues.docx");
    }
}
