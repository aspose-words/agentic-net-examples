using System;
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
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a primary series.
        string[] categories = new[] { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Primary Series", categories, new double[] { 1000, 1500, 1200, 1300 });

        // Create a secondary series group and assign it to the secondary axis.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Set the number format of the secondary Y‑axis to currency with two decimals.
        secondaryGroup.AxisY.NumberFormat.FormatCode = "\"$\"#,##0.00";

        // Optional: give the secondary Y‑axis a title.
        secondaryGroup.AxisY.Title.Show = true;
        secondaryGroup.AxisY.Title.Text = "Secondary Y Axis (USD)";

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 200, 250, 220, 210 });

        // Save the document.
        doc.Save("SetSecondaryYAxisNumberFormat.docx");
    }
}
