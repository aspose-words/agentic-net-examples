using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and clear its demo data.
        Chart chart = AppendChart(builder, ChartType.Column, 500, 300);

        // Define categories and values for a single series.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.75, 160.2 };

        // Add the series to the chart.
        ChartSeries series = chart.Series.Add("Revenue", categories, values);

        // Set the series fill color.
        series.Format.Fill.ForeColor = Color.Green;

        // Enable data labels and show the value for each point.
        series.HasDataLabels = true;
        series.DataLabels.ShowValue = true;

        // Save the document.
        doc.Save("ChartSeriesExample.docx");
    }

    // Helper method to insert a chart and clear its default series.
    private static Chart AppendChart(DocumentBuilder builder, ChartType chartType, double width, double height)
    {
        Shape chartShape = builder.InsertChart(chartType, width, height);
        Chart chart = chartShape.Chart;
        chart.Series.Clear();
        return chart;
    }
}
