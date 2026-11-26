using System;
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

        // Insert a column chart with explicit size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Define custom data: series names, categories, and values.
        string[] seriesNames = { "Series A", "Series B", "Series C" };
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[,] values = {
            { 10, 20, 30, 40 },
            { 15, 25, 35, 45 },
            { 12, 22, 32, 42 }
        };

        // Remove any default series that were added automatically.
        chart.Series.Clear();

        // Set category labels for the X‑axis if the API supports it.
        // Some Aspose.Words versions expose CategoryAxis; if not, this line can be omitted.
        // chart.CategoryAxis.CategoryLabels = categories;

        // Add each series with its corresponding values.
        for (int i = 0; i < seriesNames.Length; i++)
        {
            double[] seriesValues = new double[categories.Length];
            for (int j = 0; j < categories.Length; j++)
            {
                seriesValues[j] = values[i, j];
            }
            chart.Series.Add(seriesNames[i], seriesValues);
        }

        // Save the document containing the populated chart.
        doc.Save("CustomDataChart.docx");
    }
}
