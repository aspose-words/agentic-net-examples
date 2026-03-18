using Aspose.Words;
using Aspose.Words.Drawing; // <-- added namespace for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart (suitable for financial data).
        Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = shape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a primary series with some sample data.
        string[] categories = new[] { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Primary Series", categories, new double[] { 12000, 15000, 13000, 17000 });

        // Create a secondary series group and assign it to the secondary axis.
        ChartSeriesGroup secondaryGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
        secondaryGroup.AxisGroup = AxisGroup.Secondary;

        // Optionally hide the secondary X‑axis if it is not required.
        secondaryGroup.AxisX.Hidden = true;

        // Add a series to the secondary group.
        secondaryGroup.Series.Add("Secondary Series", categories, new double[] { 8000, 9000, 8500, 9500 });

        // Set the secondary Y‑axis number format to currency with two decimal places.
        secondaryGroup.AxisY.NumberFormat.FormatCode = "\"$\"#,##0.00";
        // Ensure the format is taken from the code, not from a source cell.
        secondaryGroup.AxisY.NumberFormat.IsLinkedToSource = false;

        // Save the document.
        doc.Save("SecondaryAxisCurrency.docx");
    }
}
