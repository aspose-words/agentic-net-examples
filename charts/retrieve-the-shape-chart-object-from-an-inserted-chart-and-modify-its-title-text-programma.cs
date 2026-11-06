using System;
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

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Modify the chart title text.
        chart.Title.Text = "Quarterly Sales";

        // Save the document with the updated chart title.
        doc.Save("ChartTitleModified.docx");
    }
}
