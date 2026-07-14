using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart related classes

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. The default chart contains three demo series.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Verify that the chart has at least three series before attempting removal.
        if (chart.Series.Count < 3)
            throw new InvalidOperationException("The chart does not contain enough series to remove.");

        // Remove the third series (zero‑based index 2).
        chart.Series.RemoveAt(2);

        // Save the modified document.
        doc.Save("RemoveSeries.docx");
    }
}
