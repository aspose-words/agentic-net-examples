using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class InsertShapeWithChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart shape (the shape itself contains a chart).
        // This uses the InsertChart method which returns a Shape object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Access the Chart object to customize it (optional).
        Chart chart = chartShape.Chart;
        chart.Title.Text = "Sample Column Chart";
        chart.Title.Show = true;
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.Blue;

        // Add a data series to the chart.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 10.0, 20.0, 30.0, 40.0 });

        // Save the document as a DOCX file.
        doc.Save("ShapeWithChart.docx");
    }
}
