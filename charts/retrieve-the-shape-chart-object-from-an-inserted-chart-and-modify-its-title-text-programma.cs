using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        // Width = 432 points, Height = 252 points.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Retrieve the Chart object from the inserted shape.
        Chart chart = chartShape.Chart;

        // Modify the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview 2024";
        title.Font.Size = 16;
        title.Font.Color = System.Drawing.Color.DarkBlue;
        title.Show = true;      // Ensure the title is visible.
        title.Overlay = false;  // Keep other elements from overlapping the title.

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedChart.docx");
        doc.Save(outputPath);
    }
}
