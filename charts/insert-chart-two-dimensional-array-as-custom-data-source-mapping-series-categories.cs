using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class InsertChartFrom2DArray
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and obtain its Chart object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Define categories (X‑axis labels).
        string[] categories = { "Category 1", "Category 2", "Category 3" };

        // Two‑dimensional array: rows = categories, columns = series.
        double[,] data = new double[,] {
            { 76.6, 64.2 }, // Category 1 values for Series 1 and Series 2
            { 82.1, 79.5 }, // Category 2 values
            { 91.6, 94.0 }  // Category 3 values
        };

        // Series names.
        string[] seriesNames = { "Series 1", "Series 2" };

        // Populate the chart using the Add(string, string[], double[]) overload.
        for (int seriesIndex = 0; seriesIndex < seriesNames.Length; seriesIndex++)
        {
            double[] values = new double[categories.Length];
            for (int catIndex = 0; catIndex < categories.Length; catIndex++)
            {
                // Extract the value for the current series and category.
                values[catIndex] = data[catIndex, seriesIndex];
            }

            // Add the series to the chart.
            chart.Series.Add(seriesNames[seriesIndex], categories, values);
        }

        // Save the document.
        doc.Save("ChartFrom2DArray.docx");
    }
}
