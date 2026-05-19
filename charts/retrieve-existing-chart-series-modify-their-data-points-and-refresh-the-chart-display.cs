using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a document with a column chart that contains the default demo data.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Save the document that holds the original chart.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document back from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Locate the first shape that actually contains a chart.
        Shape? shapeWithChart = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                         .OfType<Shape>()
                                         .FirstOrDefault(s => s.HasChart);
        if (shapeWithChart == null)
            throw new InvalidOperationException("No chart shape found in the document.");

        Chart loadedChart = shapeWithChart.Chart;

        // -----------------------------------------------------------------
        // 3. Modify existing series.
        // -----------------------------------------------------------------

        // Example 1: Replace all values of the first series with new data.
        if (loadedChart.Series.Count > 0)
        {
            ChartSeries firstSeries = loadedChart.Series[0];
            // Remove existing values but keep formatting.
            firstSeries.ClearValues();

            // The original demo series uses string categories, therefore we must add
            // X values as strings to keep the collection type consistent.
            firstSeries.Add(ChartXValue.FromString("Category 1"), ChartYValue.FromDouble(15));
            firstSeries.Add(ChartXValue.FromString("Category 2"), ChartYValue.FromDouble(25));
            firstSeries.Add(ChartXValue.FromString("Category 3"), ChartYValue.FromDouble(35));
            firstSeries.Add(ChartXValue.FromString("Category 4"), ChartYValue.FromDouble(45));
        }

        // Example 2: Change the second data point of the second series.
        if (loadedChart.Series.Count > 1)
        {
            ChartSeries secondSeries = loadedChart.Series[1];
            // Ensure the series has at least two points.
            if (secondSeries.DataPoints.Count > 1)
            {
                // Remove the existing point at index 1.
                secondSeries.Remove(1);
                // Insert a new point with the same X‑type (string) as the original series.
                secondSeries.Insert(1, ChartXValue.FromString("Category 2"), ChartYValue.FromDouble(30));
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the updated document.
        // -----------------------------------------------------------------
        string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "updated.docx");
        loadedDoc.Save(updatedPath);
    }
}
