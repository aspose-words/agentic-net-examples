using System;
using System.Drawing;
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

        // Ensure the legend is present.
        ChartLegend legend = chart.Legend;

        // The Legend class does not expose a direct Visible property in some
        // Aspose.Words versions, so we simply leave it visible (default).

        // Apply a border stroke to the legend using dynamic to stay compatible
        // with versions that may or may not expose LineFormat.
        try
        {
            dynamic dynLegend = legend;
            dynLegend.LineFormat.Width = 2.0;                     // Thickness in points
            dynLegend.LineFormat.DashStyle = DashStyle.DashDot; // Dash style
            dynLegend.LineFormat.FillFormat.ForeColor = Color.Black;
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            // If the current Aspose.Words version does not support LineFormat,
            // we silently ignore the styling step.
        }

        // Save the document.
        doc.Save("ChartWithLegendBorder.docx");
        Console.WriteLine("Document saved successfully.");
    }
}
