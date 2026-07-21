using System;
using System.IO;
using System.Drawing;
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

        // Retrieve the first two series from the chart.
        ChartSeries series1 = chart.Series[0];
        ChartSeries series2 = chart.Series[1];

        // Remove the first data point from each series.
        series1.Remove(0);
        series2.Remove(0);

        // Add a new data point with a custom category and value to each series.
        ChartXValue newCategory = ChartXValue.FromString("New Category");
        series1.Add(newCategory, ChartYValue.FromDouble(15.0));
        series2.Add(newCategory, ChartYValue.FromDouble(8.5));

        // Change the fill color of the newly added point in the first series.
        int newIndex = series1.DataPoints.Count - 1;
        series1.DataPoints[newIndex].Format.Fill.Color = Color.Green;

        // Save the document with the modified chart.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedChart.docx");
        doc.Save(outputPath);
    }
}
