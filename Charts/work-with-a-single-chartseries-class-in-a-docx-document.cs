using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class ChartSeriesExample
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series so we can start with a clean chart.
        chart.Series.Clear();

        // Define categories (X axis) and corresponding Y values for a single series.
        string[] categories = new string[] { "Q1", "Q2", "Q3", "Q4" };
        double[] values = new double[] { 1200, 1500, 1100, 1700 };

        // Add one series to the chart.
        ChartSeries series = chart.Series.Add("Revenue", categories, values);

        // Apply formatting to the whole series (line weight and fill color).
        series.Format.Stroke.Weight = 2.5;
        series.Format.Fill.Solid(Color.LightBlue);

        // Find the index of the maximum value to highlight it.
        int maxIndex = 0;
        double maxValue = values[0];
        for (int i = 1; i < values.Length; i++)
        {
            if (values[i] > maxValue)
            {
                maxValue = values[i];
                maxIndex = i;
            }
        }

        // Highlight the maximum data point with a red diamond marker.
        series.DataPoints[maxIndex].Marker.Symbol = MarkerSymbol.Diamond;
        series.DataPoints[maxIndex].Marker.Size = 12;
        series.DataPoints[maxIndex].Format.Fill.Color = Color.Red;

        // Save the document containing the chart.
        doc.Save("ChartSeriesSingle.docx");
    }
}
