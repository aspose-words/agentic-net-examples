using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing;

class InsertChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart with a size of 300x300 points.
        // The InsertChart method is provided by the DocumentBuilder class.
        Shape chartShape = builder.InsertChart(ChartType.Pie,
            ConvertUtil.PixelToPoint(300), ConvertUtil.PixelToPoint(300));

        // Access the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove any demo data series that were added by default.
        chart.Series.Clear();

        // Add a new series with categories (X values) and corresponding numeric values (Y values).
        chart.Series.Add("My Fruit",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 1.3, 2.2, 1.5 });

        // Optionally, customize the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Fruit Consumption";
        title.Font.Size = 14;
        title.Font.Color = Color.Blue;
        title.Show = true;

        // Save the document to a DOCX file.
        doc.Save("ChartInserted.docx", SaveFormat.Docx);
    }
}
