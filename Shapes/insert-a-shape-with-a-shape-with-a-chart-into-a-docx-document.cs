using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart shape (Bar chart) with a width of 400 points and a height of 300 points.
        // The InsertChart method returns a Shape that already contains a Chart.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

        // Access the Chart object to customize its appearance and data.
        Chart chart = chartShape.Chart;
        chart.Title.Text = "Sample Bar Chart";
        chart.Title.Show = true;
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = System.Drawing.Color.Blue;

        // Clear any default series and add a new series with categories and values.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 10.0, 20.0, 30.0 });

        // Save the document as a DOCX file.
        doc.Save("ShapeWithChart.docx");
    }
}
