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
        // File names for the initial and updated documents.
        string initialDocPath = "chart.docx";
        string updatedDocPath = "chart-updated.docx";

        // -------------------------------------------------
        // Step 1: Create a new document and insert a column chart with initial data.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove any default series and add a custom series.
        chart.Series.Clear();
        chart.Series.Add("Initial Series", new double[] { 1, 2, 3 });

        // Save the document that now contains the chart.
        doc.Save(initialDocPath);

        // -------------------------------------------------
        // Step 2: Load the document, locate the chart, and modify its series data.
        // -------------------------------------------------
        Document loadedDoc = new Document(initialDocPath);

        // Find the first shape that actually contains a chart.
        Shape shapeWithChart = loadedDoc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasChart);

        if (shapeWithChart == null)
            throw new InvalidOperationException("No chart shape found in the document.");

        Chart loadedChart = shapeWithChart.Chart;

        // Ensure there is at least one series to modify.
        if (loadedChart.Series.Count == 0)
            throw new InvalidOperationException("The chart does not contain any series.");

        // Retrieve the first series.
        ChartSeries series = loadedChart.Series[0];

        // Preserve the original series name.
        string seriesName = series.Name;

        // Remove the existing series.
        loadedChart.Series.RemoveAt(0);

        // Add a new series with the same name but updated values.
        loadedChart.Series.Add(seriesName, new double[] { 10, 20, 30 });

        // Save the updated document.
        loadedDoc.Save(updatedDocPath);
    }
}
