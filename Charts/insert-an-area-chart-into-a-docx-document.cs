using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added for Shape
using Aspose.Words.Drawing.Charts;   // Chart, ChartType
using System;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an Area chart. Width and height are specified in points.
        Shape chartShape = builder.InsertChart(ChartType.Area, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the automatically generated demo series.
        chart.Series.Clear();

        // Add a custom series with category labels and corresponding values.
        chart.Series.Add(
            "Quarterly Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 120.5, 150.0, 130.75, 160.2 });

        // Save the document to a DOCX file.
        doc.Save("AreaChart.docx");
    }
}
