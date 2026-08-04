using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

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

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a primary series (optional, just to have data).
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Primary Series", categories, new double[] { 10, 20, 30 });

        // Create a secondary series group and assign it to the secondary axis group.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Configure the secondary X axis.
        ChartAxis secondaryXAxis = secondaryGroup.AxisX;

        // Set display units to thousands.
        secondaryXAxis.DisplayUnit.Unit = AxisBuiltInUnit.Thousands;

        // Apply a custom number format to the axis labels.
        secondaryXAxis.NumberFormat.FormatCode = "#,##0";
        secondaryXAxis.NumberFormat.IsLinkedToSource = false;

        // (Optional) Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 15, 25, 35 });

        // Save the document.
        doc.Save("SecondaryXAxisDisplayUnits.docx");
    }
}
