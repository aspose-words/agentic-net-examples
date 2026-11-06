using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartFromArrayExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a defined size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        if (!chartShape.HasChart)
        {
            throw new InvalidOperationException("The inserted shape does not contain a chart.");
        }

        Chart chart = chartShape.Chart;

        // Clear any default series that were added automatically.
        chart.Series.Clear();

        // Define categories (X‑axis labels).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Define series names.
        string[] seriesNames = { "North Region", "South Region", "East Region" };

        // Two‑dimensional array holding the values.
        // Rows correspond to series, columns correspond to categories.
        double[,] values = new double[,]
        {
            { 12000, 15000, 13000, 17000 }, // North Region
            { 10000, 11000, 11500, 14000 }, // South Region
            {  9000,  9500, 10500, 12000 }  // East Region
        };

        // Populate the chart with data from the two‑dimensional array.
        for (int i = 0; i < seriesNames.Length; i++)
        {
            double[] seriesValues = new double[categories.Length];
            for (int j = 0; j < categories.Length; j++)
            {
                seriesValues[j] = values[i, j];
            }

            // Add the series to the chart.
            chart.Series.Add(seriesNames[i], categories, seriesValues);
        }

        // Optional: set chart title and legend position.
        chart.Title.Text = "Quarterly Sales by Region";
        chart.Title.Overlay = false;
        chart.Legend.Position = LegendPosition.Right;

        // Save the document containing the chart.
        string outputDoc = Path.Combine(Directory.GetCurrentDirectory(), "ChartFromArray.docx");
        doc.Save(outputDoc);

        // Export the same document to PDF to demonstrate rendering.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "ChartFromArray.pdf");
        doc.Save(outputPdf);
    }
}
