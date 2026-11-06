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

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            return;

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Remove any default series that were added automatically.
        chart.Series.Clear();

        // Add three series with sample data.
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", new double[] { 15, 25, 35 });
        chart.Series.Add("Series 3", new double[] { 12, 22, 32 });

        // Define the index of the series to remove (e.g., the second series).
        int removeIndex = 1;

        // Validate the index before attempting removal.
        if (removeIndex >= 0 && removeIndex < chart.Series.Count)
        {
            chart.Series.RemoveAt(removeIndex);
        }

        // Save the modified document.
        doc.Save("remove-series.docx");
    }
}
