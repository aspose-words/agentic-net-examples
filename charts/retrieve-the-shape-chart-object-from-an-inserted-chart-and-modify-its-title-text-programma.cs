using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Modify the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview";
        title.Font.Size = 16;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;      // Ensure the title is visible.
        title.Overlay = false;  // Do not allow other elements to overlap the title.

        // Save the document with the updated chart.
        doc.Save("ChartWithTitle.docx");
    }
}
