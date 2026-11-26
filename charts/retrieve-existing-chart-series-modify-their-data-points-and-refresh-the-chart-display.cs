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
        // 1. Create a document with a column chart and save it as the original file.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove any default series and add a custom one.
        chart.Series.Clear();
        chart.Series.Add("Quarterly Sales", new double[] { 12000, 15000, 18000 });

        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document, locate the chart, and modify its series.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        Shape loadedChartShape = loadedDoc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasChart);

        if (loadedChartShape == null)
            throw new InvalidOperationException("No chart shape found in the document.");

        Chart loadedChart = loadedChartShape.Chart;

        if (loadedChart.Series.Count == 0)
            throw new InvalidOperationException("The chart does not contain any series to modify.");

        // Remove the existing series.
        loadedChart.Series.RemoveAt(0);

        // Add a new series with updated data points and a new name.
        loadedChart.Series.Add("Updated Quarterly Sales", new double[] { 13000, 16000, 19000 });

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "modified.docx");
        loadedDoc.Save(modifiedPath);
    }
}
