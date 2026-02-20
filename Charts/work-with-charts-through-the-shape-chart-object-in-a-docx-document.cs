using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class ChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a specific size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Define categories for the X axis.
        string[] categories = new string[] { "Q1", "Q2", "Q3", "Q4" };

        // Add two data series.
        chart.Series.Add("Revenue", categories, new double[] { 15000, 20000, 18000, 22000 });
        chart.Series.Add("Profit", categories, new double[] { 4000, 5000, 4500, 6000 });

        // Set chart title.
        ChartTitle title = chart.Title;
        title.Text = "Annual Financial Summary";
        title.Font.Size = 16;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;
        title.Overlay = false;

        // Format the chart background.
        chart.Format.Fill.Solid(Color.LightYellow);

        // Position the legend at the top right and allow other elements to overlap it.
        ChartLegend legend = chart.Legend;
        legend.Position = LegendPosition.TopRight;
        legend.Overlay = true;
        legend.Format.Fill.Solid(Color.WhiteSmoke);

        // Hide major gridlines on the primary Y axis.
        foreach (ChartAxis axis in chart.Axes)
        {
            if (axis.Type == ChartAxisType.Value)
                axis.HasMajorGridlines = false;
        }

        // Set the style of the chart to a predefined style.
        chart.Style = ChartStyle.Shaded;

        // Save the document to disk.
        doc.Save("ChartExample.docx");
    }
}
