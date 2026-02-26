using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. Width and height are specified in points.
        double width = 400;   // chart width
        double height = 300;  // chart height
        Shape chartShape = builder.InsertChart(ChartType.Column, width, height);
        Chart chart = chartShape.Chart;

        // Remove the default demo series that Aspose.Words inserts.
        chart.Series.Clear();

        // Define categories (X‑axis) and corresponding values (Y‑axis).
        string[] categories = { "Category 1", "Category 2", "Category 3" };
        double[] values = { 76.6, 82.1, 91.6 };

        // Add a series with the defined data.
        chart.Series.Add("Series 1", categories, values);

        // Save the document containing the column chart.
        doc.Save("ColumnChart.docx");
    }
}
