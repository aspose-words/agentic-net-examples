using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Define categories with long text to demonstrate wrapping.
        string[] categories = new[]
        {
            "Very Long Category Name That Should Wrap",
            "Another Extremely Long Category Name For Testing",
            "Short"
        };

        // Add a series with values.
        ChartSeries series = chart.Series.Add("Sample Series", categories, new double[] { 10, 20, 30 });

        // Enable data labels.
        series.HasDataLabels = true;

        // Show both category name and value.
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowValue = true;

        // Align labels to the center of the data point.
        series.DataLabels.Position = ChartDataLabelPosition.Center;

        // Use a line break as separator to force multi‑line labels (enables wrapping).
        series.DataLabels.Separator = "\n";

        // Optional: set a readable font size.
        series.DataLabels.Font.Size = 10;

        // Save the document.
        doc.Save("AlignedWrappedDataLabels.docx");
    }
}
