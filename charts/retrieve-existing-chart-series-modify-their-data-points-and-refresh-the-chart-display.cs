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
        // Create a new document and insert a column chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Save the initial document.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        doc.Save(originalPath);

        // Load the document containing the chart.
        Document loadedDoc = new Document(originalPath);

        // Locate the first shape that contains a chart.
        Shape? shapeWithChart = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                         .OfType<Shape>()
                                         .FirstOrDefault(s => s.HasChart);
        if (shapeWithChart == null)
            throw new InvalidOperationException("No chart shape found in the document.");

        Chart loadedChart = shapeWithChart.Chart;

        // Modify each existing series: clear current values and add new ones.
        foreach (ChartSeries series in loadedChart.Series)
        {
            // Remove existing data while preserving formatting.
            series.ClearValues();

            // Add new data points for the same categories.
            series.Add(ChartXValue.FromString("Category 1"), ChartYValue.FromDouble(10));
            series.Add(ChartXValue.FromString("Category 2"), ChartYValue.FromDouble(20));
            series.Add(ChartXValue.FromString("Category 3"), ChartYValue.FromDouble(30));
            series.Add(ChartXValue.FromString("Category 4"), ChartYValue.FromDouble(40));
        }

        // Save the updated document.
        string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "updated.docx");
        loadedDoc.Save(updatedPath);
    }
}
