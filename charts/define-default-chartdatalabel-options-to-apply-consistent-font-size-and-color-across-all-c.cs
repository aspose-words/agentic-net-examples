using System.Drawing;
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

        // Remove the automatically generated demo series.
        chart.Series.Clear();

        // Define categories for the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add sample series with data.
        chart.Series.Add("Product A", categories, new double[] { 120, 150, 170, 130 });
        chart.Series.Add("Product B", categories, new double[] { 80, 110, 140, 100 });

        // Apply the same data label settings to every series.
        foreach (ChartSeries series in chart.Series)
        {
            series.HasDataLabels = true;               // Enable data labels.
            series.DataLabels.ShowValue = true;        // Show the numeric value.
            series.DataLabels.Font.Size = 12;          // Consistent font size.
            series.DataLabels.Font.Color = Color.DarkBlue; // Consistent font color.
        }

        // Save the document containing the chart.
        doc.Save("ChartDataLabelDefaults.docx");
    }
}
