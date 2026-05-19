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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo data.
        chart.Series.Clear();

        // Define custom category labels (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 1500, 2000, 1800, 2200 };

        // Add a new series with the custom categories and values.
        chart.Series.Add("Quarterly Sales", categories, values);

        // Save the document with the modified chart.
        doc.Save("ChartSeriesExample.docx");
    }
}
