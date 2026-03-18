using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;

class ChartSeriesModifier
{
    static void Main()
    {
        // Load an existing Word document that contains a chart.
        const string inputPath = @"ArtifactsDir\ChartDocument.docx";
        const string outputPath = @"ArtifactsDir\ChartDocument_Modified.docx";

        Document doc = new Document(inputPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Find the first shape that contains a chart.
        Shape chartShape = null;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasChart)
            {
                chartShape = shape;
                break;
            }
        }

        if (chartShape == null)
        {
            Console.WriteLine("No chart found in the document.");
            return;
        }

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Iterate through all series in the chart.
        foreach (ChartSeries series in chart.Series)
        {
            // Example modification 1: Change the first Y value (if it exists).
            if (series.YValues.Count > 0)
            {
                // Replace the first Y value with a new double value.
                series.YValues[0] = ChartYValue.FromDouble(42.0);
            }

            // Example modification 2: Add a new data point to the series.
            // Use a string category for the X value and a double for the Y value.
            ChartXValue newX = ChartXValue.FromString("New Category");
            ChartYValue newY = ChartYValue.FromDouble(15.5);
            series.Add(newX, newY);

            // Example modification 3: Remove the last data point (if more than one point exists).
            if (series.YValues.Count > 1)
            {
                int lastIndex = series.YValues.Count - 1;
                series.Remove(lastIndex);
            }
        }

        // Save the modified document. The chart reflects the changes automatically.
        doc.Save(outputPath);
    }
}
