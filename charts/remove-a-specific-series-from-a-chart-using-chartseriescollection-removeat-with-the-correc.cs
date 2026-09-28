using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class RemoveChartSeriesExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Clear any default series that Aspose.Words may have added.
        chart.Series.Clear();

        // Add three series to the chart.
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", new double[] { 15, 25, 35 });
        chart.Series.Add("Series 3", new double[] { 12, 22, 32 });

        // Index of the series to remove (e.g., remove the second series).
        int removeIndex = 1;

        // Validate the index before removal.
        if (removeIndex < 0 || removeIndex >= chart.Series.Count)
            throw new ArgumentOutOfRangeException(nameof(removeIndex), "Invalid series index.");

        // Remove the series at the specified index.
        chart.Series.RemoveAt(removeIndex);

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartSeriesRemoved.docx");
        doc.Save(outputPath);
    }
}
