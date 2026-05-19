using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace RemoveChartSeriesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart that contains the default demo series.
            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
            Chart chart = chartShape.Chart;

            // Get the series collection from the chart.
            ChartSeriesCollection series = chart.Series;

            // Index of the series we want to remove (e.g., the second series).
            int indexToRemove = 1;

            // Validate the index before calling RemoveAt.
            if (indexToRemove >= 0 && indexToRemove < series.Count)
            {
                series.RemoveAt(indexToRemove);
                Console.WriteLine($"Series at index {indexToRemove} removed.");
            }
            else
            {
                Console.WriteLine($"Index {indexToRemove} is out of range. No series removed.");
            }

            // Save the modified document.
            doc.Save("RemoveSeries.docx");
        }
    }
}
