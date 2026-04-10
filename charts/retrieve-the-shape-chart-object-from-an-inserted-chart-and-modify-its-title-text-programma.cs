using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        // Width = 400 points, Height = 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Verify that the inserted shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Retrieve the Chart object from the shape.
        Chart chart = chartShape.Chart;

        // Access the chart title and modify its text.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview Q1";
        title.Show = true; // Ensure the title is visible.

        // Optionally, customize the title font.
        title.Font.Size = 16;
        title.Font.Color = System.Drawing.Color.DarkBlue;

        // Save the document to the local file system.
        doc.Save("ChartTitleModified.docx");
    }
}
