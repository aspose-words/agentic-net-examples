using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear default demo data.
        chart.Series.Clear();

        // Add custom series data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120, 150, 180, 200 };
        chart.Series.Add("Sales", categories, values);

        // Hide the plot area border while keeping axis lines visible.
        chart.Format.Stroke.Weight = 0;
        chart.Format.Stroke.Color = Color.Transparent;

        // Save the document.
        doc.Save("HidePlotAreaBorder.docx");
    }
}
