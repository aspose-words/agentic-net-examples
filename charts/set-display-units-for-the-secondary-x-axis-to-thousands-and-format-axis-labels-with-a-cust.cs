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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a primary series (optional, just to have data).
        chart.Series.Add("Primary Series",
            new[] { "A", "B", "C" },
            new double[] { 10, 20, 30 });

        // Create a secondary series group.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Configure the secondary X‑axis: set display units to thousands.
        secondaryGroup.AxisX.DisplayUnit.Unit = AxisBuiltInUnit.Thousands;

        // Apply a custom number format to the secondary X‑axis labels.
        secondaryGroup.AxisX.NumberFormat.FormatCode = "#,##0";
        // Ensure the format is not linked to the source data.
        secondaryGroup.AxisX.NumberFormat.IsLinkedToSource = false;

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series",
            new[] { "A", "B", "C" },
            new double[] { 1000, 2000, 3000 });

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SecondaryAxisChart.docx");
        doc.Save(outputPath);
    }
}
