using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class HidePlotAreaBorder
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear the demo data series and add custom data.
        chart.Series.Clear();
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 150, 200, 180, 220 });

        // Hide the plot area border while keeping axis lines visible.
        // Set the border (stroke) weight to zero and make it transparent.
        chart.Format.Stroke.Weight = 0;
        chart.Format.Stroke.Color = Color.Transparent;

        // Save the document.
        doc.Save("HidePlotAreaBorder.docx");
    }
}
