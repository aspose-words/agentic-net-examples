using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a custom series with sample data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 1500, 2300, 1800, 2100 });

        // Hide the plot area border by making the chart's line transparent.
        chart.Format.Stroke.Color = Color.Transparent;

        // Axis lines remain visible (default behavior).

        // Save the document.
        doc.Save("HidePlotAreaBorder.docx");
    }
}
