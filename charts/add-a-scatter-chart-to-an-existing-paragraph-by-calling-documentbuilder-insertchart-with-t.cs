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

        // Add a paragraph that will contain the chart.
        builder.Writeln("Scatter chart example:");

        // Insert a scatter chart into the current paragraph.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 400, 300);
        Chart chart = chartShape.Chart;

        // Set a visible title for the chart.
        chart.Title.Text = "Sample Scatter Chart";
        chart.Title.Show = true;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add the first data series.
        chart.Series.Add(
            "Series 1",
            new double[] { 1.0, 2.5, 4.0, 5.5 },
            new double[] { 2.0, 3.5, 1.0, 4.5 });

        // Add the second data series.
        chart.Series.Add(
            "Series 2",
            new double[] { 1.5, 3.0, 4.5, 6.0 },
            new double[] { 1.5, 2.0, 3.5, 2.5 });

        // Save the document to the working directory.
        doc.Save("ScatterChart.docx");
    }
}
