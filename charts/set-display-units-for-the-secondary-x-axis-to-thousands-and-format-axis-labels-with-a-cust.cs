using System;
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
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a primary series (optional, just to have data on the primary axis).
        string[] categories = { "Jan", "Feb", "Mar" };
        chart.Series.Add("Primary Series", categories, new double[] { 10, 20, 30 });

        // Create a secondary series group that uses the secondary axes.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Set the secondary X‑axis display unit to thousands.
        secondaryGroup.AxisX.DisplayUnit.Unit = AxisBuiltInUnit.Thousands;

        // Apply a custom number format to the secondary X‑axis labels.
        secondaryGroup.AxisX.NumberFormat.FormatCode = "#,##0";

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 1000, 2000, 3000 });

        // Save the document.
        doc.Save("SecondaryXAxisDisplayUnits.docx");
    }
}
