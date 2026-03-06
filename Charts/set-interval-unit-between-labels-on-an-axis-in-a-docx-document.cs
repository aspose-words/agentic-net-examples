using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a simple series with categories and values.
        chart.Series.Add("Sample Series",
            new[] { "Category A", "Category B", "Category C", "Category D" },
            new double[] { 10, 20, 30, 40 });

        // Get the X‑axis of the chart.
        ChartAxis xAxis = chart.AxisX;

        // Turn off automatic spacing for tick labels.
        xAxis.TickLabels.IsAutoSpacing = false;

        // Set the interval (spacing) between tick labels.
        // For example, a value of 2 draws every second label.
        xAxis.TickLabels.Spacing = 2;

        // Save the document to a DOCX file.
        doc.Save("AxisLabelSpacing.docx");
    }
}
