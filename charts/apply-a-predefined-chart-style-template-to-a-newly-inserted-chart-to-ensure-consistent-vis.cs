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
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300, ChartStyle.ShadedPlot);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add custom series data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 1500, 2000, 1800, 2200 });

        // Set a visible title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document to the working directory.
        doc.Save("ChartWithStyle.docx");
    }
}
