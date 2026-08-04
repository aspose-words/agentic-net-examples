using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that comes with the chart.
        chart.Series.Clear();

        // Add a new series with placeholder categories and values.
        // These placeholders will be replaced with the actual data later.
        string[] placeholderCategories = { "Cat1", "Cat2", "Cat3" };
        double[] placeholderValues = { 0, 0, 0 };
        chart.Series.Add("Quarterly Sales", placeholderCategories, placeholderValues);

        // Retrieve the newly added series.
        ChartSeries series = chart.Series[0];

        // Clear existing values while preserving formatting.
        series.ClearValues();

        // Set the actual numeric values with custom category labels.
        series.Add(ChartXValue.FromString("Q1 2023"), ChartYValue.FromDouble(15000));
        series.Add(ChartXValue.FromString("Q2 2023"), ChartYValue.FromDouble(23000));
        series.Add(ChartXValue.FromString("Q3 2023"), ChartYValue.FromDouble(18000));

        // Save the document with the modified chart.
        doc.Save("ChartSeriesExample.docx");
    }
}
