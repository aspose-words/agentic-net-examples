using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;

class ChartSeriesCollectionDemo
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart (default demo data will be generated).
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Access the series collection of the chart.
        ChartSeriesCollection seriesCollection = chart.Series;

        // -----------------------------------------------------------------
        // 1. Enumerate existing series and output their names.
        // -----------------------------------------------------------------
        Console.WriteLine("Existing series:");
        using (IEnumerator<ChartSeries> enumerator = seriesCollection.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                Console.WriteLine("- " + enumerator.Current.Name);
            }
        }

        // -----------------------------------------------------------------
        // 2. Add a new series with custom categories and values.
        // -----------------------------------------------------------------
        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 12.5, 23.8, 31.4 };
        ChartSeries newSeries = seriesCollection.Add("Custom Series", categories, values);
        Console.WriteLine($"Added series: {newSeries.Name}");

        // -----------------------------------------------------------------
        // 3. Remove the second series (index 1) from the collection.
        // -----------------------------------------------------------------
        if (seriesCollection.Count > 1)
        {
            Console.WriteLine($"Removing series at index 1: {seriesCollection[1].Name}");
            seriesCollection.RemoveAt(1);
        }

        // -----------------------------------------------------------------
        // 4. Clear all series from the chart (optional, shows usage of Clear).
        // -----------------------------------------------------------------
        // seriesCollection.Clear();

        // Save the document to disk.
        string outputPath = "ChartSeriesCollection.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
