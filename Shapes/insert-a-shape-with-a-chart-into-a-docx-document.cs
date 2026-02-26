using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class InsertChartShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart (which is a shape that contains a chart) – a 400x300 point bar chart.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

        // Access the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Clear any default series and add a custom series.
        chart.Series.Clear();
        chart.Series.Add("Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 1500.0, 2300.0, 1800.0, 2100.0 });

        // Optionally set a title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.Blue;
        chart.Title.Show = true;

        // Save the document to a DOCX file.
        doc.Save("ChartShape.docx");
    }
}
