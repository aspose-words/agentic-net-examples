using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Hide every axis of the chart.
        foreach (ChartAxis axis in chart.Axes)
        {
            axis.Hidden = true;
        }

        // Save the resulting DOCX file.
        doc.Save("HideChartAxis.docx");
    }
}
