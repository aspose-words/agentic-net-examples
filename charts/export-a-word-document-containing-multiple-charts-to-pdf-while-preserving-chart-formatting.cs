using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first chart (Column chart).
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!columnChartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart columnChart = columnChartShape.Chart;
        // Clear default series.
        columnChart.Series.Clear();

        // Add a series with categories and values.
        columnChart.Series.Add(
            "Sales",
            new string[] { "Q1", "Q2", "Q3" },
            new double[] { 15000, 21000, 18000 });

        // Insert a paragraph break between charts.
        builder.Writeln();

        // Insert second chart (Pie chart).
        Shape pieChartShape = builder.InsertChart(ChartType.Pie, 432, 252);
        if (!pieChartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart pieChart = pieChartShape.Chart;
        // Clear default series.
        pieChart.Series.Clear();

        // Add a series with categories and values.
        pieChart.Series.Add(
            "Revenue Share",
            new string[] { "Product A", "Product B", "Product C" },
            new double[] { 45, 30, 25 });

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document with charts.
        string docxPath = Path.Combine(outputDir, "ChartsDocument.docx");
        doc.Save(docxPath);

        // Export the same document to PDF, preserving chart formatting.
        string pdfPath = Path.Combine(outputDir, "ChartsDocument.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
