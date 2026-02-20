using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bar chart with the specified size.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
        Chart chart = chartShape.Chart;

        // Set the chart title.
        ChartTitle title = chart.Title;
        title.Text = "Sales Report 2025";
        title.Font.Size = 16;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;
        title.Overlay = false;

        // Apply a predefined chart style.
        chart.Style = ChartStyle.Gradient;

        // Position the legend at the top right and allow other elements to overlap it.
        ChartLegend legend = chart.Legend;
        legend.Position = LegendPosition.TopRight;
        legend.Overlay = true;

        // Fill the chart background with a solid color.
        chart.Format.Fill.Solid(Color.LightYellow);

        // Hide tick labels on both axes.
        chart.AxisX.TickLabels.Position = AxisTickLabelPosition.None;
        chart.AxisY.TickLabels.Position = AxisTickLabelPosition.None;

        // Save the document to a DOCX file.
        doc.Save("ChartExample.docx");
    }
}
