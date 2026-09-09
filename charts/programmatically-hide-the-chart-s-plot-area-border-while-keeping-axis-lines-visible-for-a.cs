using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Sample Series", categories, new double[] { 10, 20, 30 });

        // Hide the plot area border while keeping axis lines visible.
        chart.Format.Stroke.Weight = 0;

        // Save the document.
        doc.Save("HidePlotAreaBorder.docx");
    }
}
