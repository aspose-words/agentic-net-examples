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

        // Verify that the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
        {
            throw new InvalidOperationException("The inserted shape does not contain a chart.");
        }

        // Retrieve the Chart object from the shape.
        Chart chart = chartShape.Chart;

        // Modify the chart title.
        chart.Title.Text = "Sales Overview";
        chart.Title.Show = true;

        // Save the document with the updated chart.
        doc.Save("chart-with-title.docx");
    }
}
