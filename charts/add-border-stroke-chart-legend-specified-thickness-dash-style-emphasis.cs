using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ChartLegendBorderExample
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Apply a border (stroke) to the chart legend.
        // Set the thickness of the border.
        chart.Legend.Format.Stroke.Weight = 2.0; // points

        // Set the dash style of the border.
        chart.Legend.Format.Stroke.DashStyle = DashStyle.Dash;

        // Optionally set the border color.
        chart.Legend.Format.Stroke.Color = Color.DarkBlue;

        // Save the document.
        doc.Save("ChartLegendBorder.docx");
    }
}
