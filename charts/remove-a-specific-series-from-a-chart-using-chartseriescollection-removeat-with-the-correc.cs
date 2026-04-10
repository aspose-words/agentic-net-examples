using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. The default chart contains three demo series.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Verify that the inserted shape actually has a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Choose the index of the series to remove.
        // Here we remove the second series (zero‑based index 1) if it exists.
        int removeIndex = 1;
        if (removeIndex < 0 || removeIndex >= chart.Series.Count)
            throw new ArgumentOutOfRangeException(nameof(removeIndex), "Series index is out of range.");

        // Remove the series at the specified index.
        chart.Series.RemoveAt(removeIndex);

        // Save the resulting document.
        doc.Save("RemoveSeriesChart.docx");
    }
}
