using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Remove any default series.
        chart.Series.Clear();

        // Add a new series with the desired values.
        double[] values = new double[] { 15, 30, 45 };
        ChartSeries series = chart.Series.Add("Sample Series", values);

        // Set fill colors for each data point.
        if (series.DataPoints.Count >= 3)
        {
            series.DataPoints[0].Format.Fill.ForeColor = Color.Red;
            series.DataPoints[1].Format.Fill.ForeColor = Color.Green;
            series.DataPoints[2].Format.Fill.ForeColor = Color.Blue;
        }

        // Save the document.
        doc.Save("ChartDataPoints.docx");
    }
}
