using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // User‑defined preferences.
        string newTitleText = "Sales Overview";
        bool legendShouldBeVisible = false; // toggle legend visibility

        // Update the chart title.
        ChartTitle title = chart.Title;
        title.Text = newTitleText;
        title.Show = true; // ensure the title is displayed.

        // Toggle the legend visibility.
        ChartLegend legend = chart.Legend;
        legend.Position = legendShouldBeVisible
            ? LegendPosition.Right   // show legend on the right.
            : LegendPosition.None;   // hide the legend.

        // Save the modified document.
        doc.Save("UpdatedChart.docx");
    }
}
