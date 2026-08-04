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

        // Insert a column chart with the default demo data.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        // Access the chart object (optional, shown for completeness).
        Chart chart = chartShape.Chart;

        // Save the document to the current working directory.
        doc.Save("insert-chart.docx");
    }
}
