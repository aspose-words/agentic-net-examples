using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart of size 500x300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that come with the chart.
        chart.Series.Clear();

        // Define category labels for the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add two new series with values for each category.
        chart.Series.Add("Revenue", categories, new double[] { 120.5, 150.2, 130.0, 170.3 });
        chart.Series.Add("Profit",  categories, new double[] { 30.1, 45.3, 35.6, 55.2 });

        // Iterate over the series collection and output each series name.
        foreach (ChartSeries series in chart.Series)
        {
            Console.WriteLine($"Series name: {series.Name}");
        }

        // Change the fill colour of the series.
        chart.Series[0].Format.Fill.ForeColor = Color.Blue;   // Revenue series
        chart.Series[1].Format.Fill.ForeColor = Color.Green;  // Profit series

        // Remove the second series (Profit) by its index.
        chart.Series.RemoveAt(1);

        // Save the document containing the modified chart.
        doc.Save("ChartSeriesDemo.docx");
    }
}
