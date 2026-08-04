using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart‑related APIs

public class RemoveChartSeriesExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. The default chart contains three demo series.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Verify that the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the series collection of the chart.
        ChartSeriesCollection series = chart.Series;

        // Define the zero‑based index of the series we want to remove.
        int indexToRemove = 2; // Removes the third demo series.

        // Validate the index before attempting removal.
        if (indexToRemove < 0 || indexToRemove >= series.Count)
            throw new ArgumentOutOfRangeException(nameof(indexToRemove), "Series index is out of range.");

        // Remove the series at the specified index.
        series.RemoveAt(indexToRemove);

        // Save the modified document.
        doc.Save("RemoveSeries.docx");
    }
}
