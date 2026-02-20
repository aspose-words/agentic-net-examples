using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

namespace AsposeWordsChartSeriesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with a predefined size.
            Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = chartShape.Chart;

            // The chart comes with demo data (three series). Remove it to start fresh.
            chart.Series.Clear();

            // Define categories for the X axis.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };

            // Add two new series with values for each category.
            chart.Series.Add("Revenue", categories, new double[] { 12000, 15000, 13000, 17000 });
            chart.Series.Add("Expenses", categories, new double[] { 8000, 9000, 8500, 9500 });

            // Enumerate the series collection and output each series name.
            foreach (ChartSeries series in chart.Series)
            {
                Console.WriteLine($"Series name: {series.Name}");
            }

            // Add a third series.
            chart.Series.Add("Profit", categories, new double[] { 4000, 6000, 4500, 7500 });

            // Remove the second series (index 1, which is "Expenses").
            chart.Series.RemoveAt(1);

            // Verify the removal.
            Console.WriteLine("\nAfter removal:");
            foreach (ChartSeries series in chart.Series)
            {
                Console.WriteLine($"Series name: {series.Name}");
            }

            // Save the document to disk.
            doc.Save("ChartSeriesCollectionDemo.docx");
        }
    }
}
