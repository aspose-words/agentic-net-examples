using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series to start with a clean chart.
        chart.Series.Clear();

        // Define category labels and corresponding numeric values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.75, 160.2 };

        // Add a new series with a name, categories, and values in one step.
        chart.Series.Add("Sales 2023", categories, values);

        // Save the document containing the chart.
        doc.Save("ChartSeriesLabeled.docx");
    }
}
