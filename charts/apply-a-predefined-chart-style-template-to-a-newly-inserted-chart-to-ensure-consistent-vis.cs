using System;
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

        // Insert a column chart with a predefined style (ShadedPlot) to ensure consistent branding.
        // Width and height are specified in points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300, ChartStyle.ShadedPlot);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add custom data series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 15000, 21000, 18000, 24000 };
        chart.Series.Add("Quarterly Sales", categories, values);

        // Optionally, set a title for the chart.
        chart.Title.Text = "Sales Overview";
        chart.Title.Show = true;

        // Save the document to the working directory.
        doc.Save("ChartWithStyle.docx");
    }
}
