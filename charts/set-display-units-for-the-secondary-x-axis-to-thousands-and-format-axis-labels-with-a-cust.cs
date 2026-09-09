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

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a primary series with sample data.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Primary Series", categories, new double[] { 10, 20, 30 });

        // Create a secondary series group.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Set the secondary X‑axis display units to thousands.
        secondaryGroup.AxisX.DisplayUnit.Unit = AxisBuiltInUnit.Thousands;

        // Apply a custom number format to the secondary X‑axis labels.
        secondaryGroup.AxisX.NumberFormat.FormatCode = "#,##0.0";

        // Add a series to the secondary group (optional).
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 15, 25, 35 });

        // Save the document.
        doc.Save("SecondaryAxisDisplayUnits.docx");
    }
}
