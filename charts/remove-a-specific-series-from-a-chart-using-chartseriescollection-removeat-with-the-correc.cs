using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;        // Chart APIs

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

        // Index of the series to remove (e.g., the second series).
        int seriesIndexToRemove = 1;

        // Validate the index before removing the series.
        if (seriesIndexToRemove >= 0 && seriesIndexToRemove < chart.Series.Count)
        {
            chart.Series.RemoveAt(seriesIndexToRemove);
        }

        // Save the modified document.
        doc.Save("RemoveSeries.docx");
    }
}
