using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;        // Chart related types

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart that will be used for financial data.
        Shape chartShape = builder.InsertChart(ChartType.Line, 450, 250);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a primary series (uses the primary Y‑axis).
        string[] categories = new[] { "Q1", "Q2", "Q3" };
        chart.Series.Add("Primary Series", categories, new double[] { 2000, 3500, 2800 });

        // Create a secondary series group that uses the secondary axes.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Optionally hide the secondary X‑axis (not required for the task).
        secondaryGroup.AxisX.Hidden = true;

        // Set the number format of the secondary Y‑axis to currency with two decimal places.
        secondaryGroup.AxisY.NumberFormat.FormatCode = "\"$\"#,##0.00";

        // Add a series that will be plotted against the secondary Y‑axis.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 1500, 2200, 3100 });

        // Save the document.
        doc.Save("SetSecondaryYAxisNumberFormat.docx");
    }
}
