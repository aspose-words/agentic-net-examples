using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a new document and insert a column chart with default demo data.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Save the original document (optional, just to show the starting point).
        string originalPath = Path.Combine(outputDir, "OriginalChart.docx");
        doc.Save(originalPath);

        // 2. Retrieve the chart shape and validate that it really contains a chart.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        Shape? targetShape = null;
        foreach (Shape shape in shapes)
        {
            if (shape.HasChart)
            {
                targetShape = shape;
                break;
            }
        }

        if (targetShape == null)
            throw new InvalidOperationException("No chart shape found in the document.");

        Chart targetChart = targetShape.Chart;

        // 3. Modify each existing series' data points.
        // Define new categories (must match the number of points we will add).
        string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };

        // Iterate over all series in the chart.
        for (int seriesIndex = 0; seriesIndex < targetChart.Series.Count; seriesIndex++)
        {
            ChartSeries series = targetChart.Series[seriesIndex];

            // Clear existing values while preserving formatting.
            series.ClearValues();

            // Create new Y values for this series (example: base value + offset per series).
            double[] newValues = new double[categories.Length];
            for (int i = 0; i < categories.Length; i++)
                newValues[i] = 10 + seriesIndex * 5 + i * 2; // arbitrary demo data

            // Populate the series with the new category/value pairs.
            for (int i = 0; i < categories.Length; i++)
            {
                series.Add(
                    ChartXValue.FromString(categories[i]),
                    ChartYValue.FromDouble(newValues[i]));
            }

            // Optionally rename the series to reflect the change.
            series.Name = $"Modified Series {seriesIndex + 1}";
        }

        // 4. Save the modified document.
        string modifiedPath = Path.Combine(outputDir, "ModifiedChart.docx");
        doc.Save(modifiedPath);
    }
}
