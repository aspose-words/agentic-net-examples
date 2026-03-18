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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and apply the predefined "Blue" style.
        // This uses the InsertChart overload that accepts a ChartStyle parameter.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300, ChartStyle.Blue);
        Chart chart = chartShape.Chart;

        // Optional: set a title for the chart.
        chart.Title.Text = "Sales Overview";
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.White;
        chart.Title.Show = true;

        // Save the document with the styled chart.
        doc.Save("ChartWithStyle.docx");
    }
}
