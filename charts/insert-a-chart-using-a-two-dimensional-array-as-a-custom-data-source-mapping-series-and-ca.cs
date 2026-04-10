using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words adds.
        chart.Series.Clear();

        // Define the categories that will appear on the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Define the names of the series (each row in the data array).
        string[] seriesNames = { "Product A", "Product B", "Product C" };

        // Two‑dimensional array of values: rows = series, columns = categories.
        double[,] values = {
            { 120.5, 135.0, 150.2, 160.8 },
            {  80.3,  95.6, 100.1, 110.4 },
            { 200.0, 210.5, 190.3, 205.7 }
        };

        // Add each series to the chart using the appropriate overload.
        for (int i = 0; i < seriesNames.Length; i++)
        {
            double[] rowValues = new double[categories.Length];
            for (int j = 0; j < categories.Length; j++)
            {
                rowValues[j] = values[i, j];
            }

            chart.Series.Add(seriesNames[i], categories, rowValues);
        }

        // Set a visible title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Save the document to the working directory.
        doc.Save("CustomChart.docx");
    }
}
