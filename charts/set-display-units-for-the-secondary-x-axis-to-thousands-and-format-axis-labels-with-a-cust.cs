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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a primary series.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Primary Series", categories, new double[] { 10, 20, 30 });

        // Create a secondary series group and assign it to the secondary axis.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Column);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 15, 25, 35 });

        // Configure the secondary X‑axis: set display units to thousands and apply a custom number format.
        ChartAxis secondaryXAxis = secondaryGroup.AxisX;
        secondaryXAxis.DisplayUnit.Unit = AxisBuiltInUnit.Thousands;
        secondaryXAxis.NumberFormat.FormatCode = "#,##0";
        secondaryXAxis.NumberFormat.IsLinkedToSource = false;

        // Save the document.
        doc.Save("ChartSecondaryXAxis.docx");
    }
}
