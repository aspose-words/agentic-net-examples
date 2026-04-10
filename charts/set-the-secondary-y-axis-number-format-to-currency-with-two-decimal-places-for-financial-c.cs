using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace AsposeChartsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart into the document.
            Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);

            // Ensure the shape actually contains a chart.
            if (!chartShape.HasChart)
                throw new InvalidOperationException("The inserted shape does not contain a chart.");

            Chart chart = chartShape.Chart;

            // Remove the demo data that Aspose.Words adds by default.
            chart.Series.Clear();

            // Primary series (uses the primary Y‑axis).
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            chart.Series.Add("Primary Series", categories, new double[] { 1000, 1500, 2000, 2500 });

            // Create a secondary series group that will use the secondary axes.
            ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Column);
            secondaryGroup.AxisGroup = AxisGroup.Secondary;          // Use secondary axes.
            secondaryGroup.AxisX.Hidden = true;                     // Hide the secondary X‑axis (optional).

            // Add a series to the secondary group (uses the secondary Y‑axis).
            ChartSeries secondarySeries = secondaryGroup.Series.Add(
                "Secondary Series", categories, new double[] { 3000, 3500, 4000, 4500 });

            // Set the number format of the secondary Y‑axis to currency with two decimal places.
            secondaryGroup.AxisY.NumberFormat.FormatCode = "\"$\"#,##0.00";

            // Save the document.
            doc.Save("SecondaryYAxisNumberFormat.docx");
        }
    }
}
