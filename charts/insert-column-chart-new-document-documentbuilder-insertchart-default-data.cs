using System;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with default demo data.
        // Width and height are specified in points (1 point = 1/72 inch).
        // Here we use 400 points wide and 300 points high.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // The chart already contains sample series; no further data manipulation is required.

        // Save the document to a file.
        doc.Save("ColumnChart.docx");
    }
}
