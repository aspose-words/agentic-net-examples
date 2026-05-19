using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a primary series (uses the primary Y‑axis).
        chart.Series.Add("Primary Series",
            new[] { "Q1", "Q2", "Q3" },
            new double[] { 1200, 1500, 1800 });

        // Create a secondary series group and assign it to the secondary axis group.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Add a series to the secondary group (will be plotted against the secondary Y‑axis).
        secondaryGroup.Series.Add("Secondary Series",
            new[] { "Q1", "Q2", "Q3" },
            new double[] { 2000, 2500, 3000 });

        // Set the number format of the secondary Y‑axis to currency with two decimal places.
        secondaryGroup.AxisY.NumberFormat.FormatCode = "\"$\"#,##0.00";
        secondaryGroup.AxisY.NumberFormat.IsLinkedToSource = false;

        // Save the document containing the chart.
        doc.Save("SecondaryYAxisCurrency.docx");
    }
}
