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

        // Insert a line chart.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Ensure we have a clean chart.
        chart.Series.Clear();

        // Add a primary series (Revenue).
        ChartSeries revenueSeries = chart.Series.Add("Revenue", new double[] { 120000, 150000, 130000, 170000 });

        // Add a secondary series (Profit). Aspose.Words does not expose a direct
        // PlotOnSecondAxis property, so both series will be plotted on the primary Y‑axis.
        ChartSeries profitSeries = chart.Series.Add("Profit", new double[] { 30000, 40000, 35000, 45000 });

        // Set the primary Y‑axis number format to currency with two decimal places.
        // The NumberFormat property is read‑only, but its FormatCode can be modified.
        ChartAxis primaryValueAxis = chart.Axes[1]; // Index 1 corresponds to the value (Y) axis.
        primaryValueAxis.NumberFormat.FormatCode = "$#,##0.00";

        // Optionally set titles for clarity.
        chart.Title.Text = "Financial Overview";
        chart.Axes[0].Title.Text = "Quarter";          // Category (X) axis
        primaryValueAxis.Title.Text = "Revenue / Profit (USD)";

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FinancialChart.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
