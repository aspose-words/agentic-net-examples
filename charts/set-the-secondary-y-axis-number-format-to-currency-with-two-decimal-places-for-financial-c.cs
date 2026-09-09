using System;
using System.IO;
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

        // Insert a line chart.
        Shape chartShape = builder.InsertChart(ChartType.Line, 450, 250);
        Chart chart = chartShape.Chart;

        // Remove the demo series.
        chart.Series.Clear();

        // Add a primary series with sample data.
        string[] categories = new[] { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Primary Series", categories, new double[] { 1200, 1500, 1800, 2100 });

        // Create a secondary series group.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Hide the secondary X axis (optional).
        secondaryGroup.AxisX.Hidden = true;

        // Set a title for the secondary Y axis.
        secondaryGroup.AxisY.Title.Show = true;
        secondaryGroup.AxisY.Title.Text = "Revenue (USD)";

        // Set the secondary Y‑axis number format to currency with two decimal places.
        secondaryGroup.AxisY.NumberFormat.FormatCode = "\"$\"#,##0.00";
        secondaryGroup.AxisY.NumberFormat.IsLinkedToSource = false;

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 3000, 3500, 4000, 4500 });

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FinancialChart.docx");
        doc.Save(outputPath);
    }
}
