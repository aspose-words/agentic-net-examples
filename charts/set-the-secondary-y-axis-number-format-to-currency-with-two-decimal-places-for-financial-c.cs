using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;        // Chart‑related types

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Primary series (uses the primary Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Primary Series", categories, new double[] { 1200, 1500, 1800, 2100 });

        // Create a secondary series group that uses the secondary axes.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Hide the secondary X‑axis (optional, keeps the chart tidy).
        secondaryGroup.AxisX.Hidden = true;

        // Set the number format of the secondary Y‑axis to currency with two decimal places.
        // Format code: "$#,##0.00"
        secondaryGroup.AxisY.NumberFormat.FormatCode = "$#,##0.00";

        // Add a series to the secondary group (uses the secondary Y‑axis).
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 3000, 3500, 4000, 4500 });

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SecondaryAxisNumberFormat.docx");
        doc.Save(outputPath);
    }
}
