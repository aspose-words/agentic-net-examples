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
        // Step 1: Create a new document and insert a sample column chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Save the document that contains the initial chart.
        const string inputPath = "chart_input.docx";
        doc.Save(inputPath);

        // Step 2: Load the document that contains the chart.
        Document loadedDoc = new Document(inputPath);
        Shape? shapeWithChart = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                         .OfType<Shape>()
                                         .FirstOrDefault(s => s.HasChart);
        if (shapeWithChart == null)
            throw new InvalidOperationException("No chart shape found in the document.");

        Chart loadedChart = shapeWithChart.Chart;

        // Step 3: Retrieve existing series and modify their data points.
        // The default chart uses string categories for the X‑axis, so we must use
        // ChartXValue.FromString when adding new points after clearing values.

        if (loadedChart.Series.Count > 0)
        {
            ChartSeries firstSeries = loadedChart.Series[0];
            firstSeries.ClearValues();

            // Add new data points with string X values (categories).
            firstSeries.Add(ChartXValue.FromString("Category 1"), ChartYValue.FromDouble(15));
            firstSeries.Add(ChartXValue.FromString("Category 2"), ChartYValue.FromDouble(30));
            firstSeries.Add(ChartXValue.FromString("Category 3"), ChartYValue.FromDouble(45));
        }

        if (loadedChart.Series.Count > 1)
        {
            ChartSeries secondSeries = loadedChart.Series[1];
            // Remove the first data point if it exists.
            if (secondSeries.YValues.Count > 0)
                secondSeries.Remove(0);

            // Append a new data point using a string category.
            secondSeries.Add(ChartXValue.FromString("Category 4"), ChartYValue.FromDouble(25));
        }

        // Step 4: Save the modified document.
        const string outputPath = "chart_modified.docx";
        loadedDoc.Save(outputPath);
    }
}
