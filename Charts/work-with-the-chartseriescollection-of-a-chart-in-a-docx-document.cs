using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with a newly inserted chart.
        chart.Series.Clear();

        // Define categories for the X‑axis.
        string[] categories = { "Category 1", "Category 2", "Category 3" };

        // Add two series with values for each category.
        chart.Series.Add("Series A", categories, new double[] { 10.5, 20.3, 30.7 });
        chart.Series.Add("Series B", categories, new double[] { 15.2, 25.8, 35.1 });

        // Iterate over the series collection using the enumerator.
        using (IEnumerator<ChartSeries> enumerator = chart.Series.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                ChartSeries current = enumerator.Current;
                Console.WriteLine($"Series name: {current.Name}");
            }
        }

        // Access a series by index (zero‑based). Change its fill colour.
        ChartSeries firstSeries = chart.Series[0];
        firstSeries.Format.Fill.ForeColor = Color.Red;

        // Remove the second series by index.
        chart.Series.RemoveAt(1);

        // Save the document to the file system.
        string artifactsDir = "output/";
        doc.Save(artifactsDir + "ChartSeriesCollectionExample.docx");
    }
}
