using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartSeriesModifier
{
    static void Main()
    {
        // Create a new document and insert a simple chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with some initial data.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Populate the chart with a single series and a few data points.
        ChartSeries series = chart.Series[0];
        series.Name = "Series 1";
        series.Add(ChartXValue.FromString("Category A"), ChartYValue.FromDouble(10));
        series.Add(ChartXValue.FromString("Category B"), ChartYValue.FromDouble(20));
        series.Add(ChartXValue.FromString("Category C"), ChartYValue.FromDouble(30));

        // Iterate through all series in the chart and modify them.
        foreach (ChartSeries s in chart.Series)
        {
            // Change the first Y value (if it exists).
            if (s.YValues.Count > 0)
                s.YValues[0] = ChartYValue.FromDouble(42.0);

            // Add a new data point.
            ChartXValue newX = ChartXValue.FromString("New Category");
            ChartYValue newY = ChartYValue.FromDouble(15.5);
            s.Add(newX, newY);

            // Remove the last data point (if more than one point exists).
            if (s.YValues.Count > 1)
            {
                int lastIndex = s.YValues.Count - 1;
                s.Remove(lastIndex);
            }
        }

        // Save the modified document.
        const string outputPath = "ChartDocument_Modified.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
