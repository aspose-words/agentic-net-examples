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

        // Verify that the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Retrieve the Chart object from the shape.
        Chart chart = chartShape.Chart;

        // Modify the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview";
        title.Show = true;               // Ensure the title is visible.
        title.Font.Size = 16;            // Set title font size.
        title.Font.Color = Color.DarkBlue; // Set title font color.

        // Save the document with the modified chart.
        doc.Save("ChartWithTitle.docx");
    }
}
