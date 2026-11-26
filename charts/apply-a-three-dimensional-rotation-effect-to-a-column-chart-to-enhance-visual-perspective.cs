using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

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

        // Aspose.Words does not expose direct 3‑D rotation properties.
        // The chart will be rendered with its default 3‑D perspective.

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Q1 Sales", new double[] { 12000, 15000, 18000 });
        chart.Series.Add("Q2 Sales", new double[] { 14000, 16000, 19000 });

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ColumnChart3D.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
