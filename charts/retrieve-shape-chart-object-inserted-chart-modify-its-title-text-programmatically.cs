using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bar chart shape with a width of 400 points and a height of 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Access the chart's title object.
        ChartTitle title = chart.Title;

        // Set the title text and customize its appearance.
        title.Text = "Sales Overview";
        title.Font.Size = 16;
        title.Font.Color = Color.DarkBlue;

        // Ensure the title is displayed.
        title.Show = true;

        // Save the document to a file.
        doc.Save("ChartWithTitle.docx");
    }
}
