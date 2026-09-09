using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Required for Shape
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 12.5, 23.8, 9.4 };
        chart.Series.Add("Sample Series", categories, values);

        // Enable data labels for the series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;

        // Align all data labels to the center of the data marker.
        series.DataLabels.Position = ChartDataLabelPosition.Center;

        // Show category name, series name and value to create multi‑line labels.
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowSeriesName = true;
        series.DataLabels.ShowValue = true;

        // Use a line break as the separator so each piece appears on a new line (wrapping).
        series.DataLabels.Separator = "\n";

        // Optional: adjust font size to improve readability.
        series.DataLabels.Font.Size = 9;
        series.DataLabels.Font.Color = Color.Black;

        // Ensure each individual label also uses the centered position.
        for (int i = 0; i < series.DataLabels.Count; i++)
        {
            series.DataLabels[i].Position = ChartDataLabelPosition.Center;
            series.DataLabels[i].Separator = "\n";
        }

        // Save the document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        doc.Save(Path.Combine(outputDir, "ChartDataLabels.docx"));
    }
}
