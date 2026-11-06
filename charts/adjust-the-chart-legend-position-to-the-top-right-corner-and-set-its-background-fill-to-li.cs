using System;
using System.Drawing;
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

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Ensure the chart has a legend and modify it.
        if (chart.Legend != null)
        {
            // Position the legend at the top right corner.
            chart.Legend.Position = LegendPosition.TopRight;

            // Set the legend background fill to light gray.
            chart.Legend.Format.Fill.Visible = true;
            chart.Legend.Format.Fill.ForeColor = Color.LightGray;
            chart.Legend.Format.Fill.Transparency = 0;
        }

        // Save the document.
        doc.Save("ChartLegendModified.docx");
    }
}
