using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Add three series to the chart.
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", new double[] { 15, 25, 35 });
        chart.Series.Add("Series 3", new double[] { 12, 22, 32 });

        // Index of the series to remove (e.g., remove the second series).
        int removeIndex = 1;

        // Validate the index before removing.
        if (removeIndex >= 0 && removeIndex < chart.Series.Count)
        {
            chart.Series.RemoveAt(removeIndex);
        }
        else
        {
            throw new ArgumentOutOfRangeException(nameof(removeIndex), "Invalid series index.");
        }

        // Save the document with the modified chart.
        doc.Save("chart-removed-series.docx");
    }
}
