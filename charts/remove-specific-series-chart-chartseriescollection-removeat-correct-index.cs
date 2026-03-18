using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class RemoveChartSeriesExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. By default it contains three demo series.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Verify the initial number of series (should be 3).
        Console.WriteLine($"Initial series count: {chart.Series.Count}");

        // Choose the index of the series to remove.
        // For example, remove the second series (zero‑based index = 1).
        int indexToRemove = 1;

        // Ensure the index is within bounds before removal.
        if (indexToRemove >= 0 && indexToRemove < chart.Series.Count)
        {
            // Remove the series at the specified index.
            chart.Series.RemoveAt(indexToRemove);
            Console.WriteLine($"Removed series at index {indexToRemove}.");
        }
        else
        {
            Console.WriteLine("Index out of range; no series removed.");
        }

        // Verify the series count after removal.
        Console.WriteLine($"Series count after removal: {chart.Series.Count}");

        // Save the document to a file.
        doc.Save("RemovedSeriesChart.docx");
    }
}
