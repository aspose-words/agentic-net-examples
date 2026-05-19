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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the scatter chart.
        builder.Writeln("Paragraph with a scatter chart:");

        // Move the builder back to the first paragraph (the one we just created).
        Paragraph paragraph = doc.FirstSection.Body.Paragraphs[0];
        builder.MoveTo(paragraph);

        // Insert a scatter chart into the paragraph using the InsertChart overload that takes ChartType, width and height.
        Shape chartShape = builder.InsertChart(ChartType.Scatter, 400, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a custom series with X and Y values.
        double[] xValues = { 1.0, 2.0, 3.0, 4.0 };
        double[] yValues = { 10.0, 20.0, 15.0, 30.0 };
        chart.Series.Add("Sample Series", xValues, yValues);

        // Save the document to the working directory.
        doc.Save("ScatterChart.docx");
    }
}
