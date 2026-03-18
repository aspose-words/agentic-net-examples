using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart (width: 500, height: 300).
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Define custom category labels (X‑axis) and corresponding numeric values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 98.3, 143.2, 110.0 };

        // Add a new series using the overload that accepts categories and values.
        ChartSeries series = chart.Series.Add("Revenue", categories, values);

        // Example of updating the series values after creation via the YValues collection.
        // Here we increase each value by 10 %.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            double original = series.YValues[i].DoubleValue;
            series.YValues[i] = ChartYValue.FromDouble(original * 1.10);
        }

        // Save the document containing the chart.
        doc.Save("ChartSeriesWithCustomCategories.docx");
    }
}
