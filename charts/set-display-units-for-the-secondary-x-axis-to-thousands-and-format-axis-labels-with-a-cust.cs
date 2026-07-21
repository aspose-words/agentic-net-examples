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
        Chart chart = chartShape.Chart;

        // Clear the default demo series.
        chart.Series.Clear();

        // Add a primary series.
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        chart.Series.Add("Primary Series", categories, new double[] { 10, 20, 30 });

        // Create a secondary series group.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary; // Use secondary axes.

        // Configure the secondary X axis.
        ChartAxis secondaryXAxis = secondaryGroup.AxisX;
        secondaryXAxis.DisplayUnit.Unit = AxisBuiltInUnit.Thousands; // Display units in thousands.
        secondaryXAxis.NumberFormat.FormatCode = "#,##0"; // Custom number format.

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 1000, 2000, 3000 });

        // Save the document.
        doc.Save("SecondaryAxisDisplayUnits.docx");
    }
}
