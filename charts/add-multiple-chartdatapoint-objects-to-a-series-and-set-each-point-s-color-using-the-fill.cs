using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart.
        Chart chart = chartShape.Chart;

        // Remove any default series.
        chart.Series.Clear();

        // Add a new series with three data points.
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });

        // Retrieve the series we just added.
        ChartSeries series = chart.Series[0];

        // Define colors for each data point.
        Color[] pointColors = new Color[] { Color.Red, Color.Green, Color.Blue };

        // Apply colors to each data point using the Fill property of the point's format.
        for (int i = 0; i < series.DataPoints.Count && i < pointColors.Length; i++)
        {
            ChartDataPoint point = series.DataPoints[i];
            point.Format.Fill.Visible = true;
            point.Format.Fill.ForeColor = pointColors[i];
        }

        // Save the document.
        doc.Save("ChartDataPointColors.docx");
    }
}
