using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart as a shape. This uses the InsertChart rule.
        // Chart type: Bar, width: 400 points, height: 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Remove any default series and add a custom series.
        chart.Series.Clear();
        chart.Series.Add(
            "Sales",                                 // Series name
            new[] { "Q1", "Q2", "Q3", "Q4" },   // Category labels
            new double[] { 15000, 20000, 18000, 22000 } // Values (double[] required)
        );

        // Set a title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document as a DOCX file.
        doc.Save("ShapeWithChart.docx");
    }
}
