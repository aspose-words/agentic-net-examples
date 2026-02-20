using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bar chart with a width of 400 points and a height of 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
        Chart chart = chartShape.Chart;

        // Set the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview";
        title.Font.Size = 14;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;

        // Move the legend to the right side of the chart.
        chart.Legend.Position = LegendPosition.Right;

        // Save the document as a DOCX file.
        doc.Save("ChartDocument.docx");
    }
}
