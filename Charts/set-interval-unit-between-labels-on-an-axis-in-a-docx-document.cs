using System;
using Aspose.Words;
using Aspose.Words.Drawing; // Added for Shape
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

class SetAxisLabelInterval
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Clear demo data and add custom series.
        chart.Series.Clear();
        chart.Series.Add("Sample Series",
            new[] { "A", "B", "C", "D" },
            new double[] { 10, 20, 30, 40 });

        // Access the X axis.
        ChartAxis xAxis = chart.AxisX;

        // Set the interval (spacing) between tick labels.
        // For example, draw a label every 2 categories.
        xAxis.TickLabels.Spacing = 2;

        // Save the document.
        doc.Save("AxisLabelInterval.docx");
    }
}
