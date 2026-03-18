using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class EnableLeaderLinesForPieChart
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart of size 500x300 points.
        Chart chart = builder.InsertChart(ChartType.Pie, 500, 300).Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        ChartSeries series = chart.Series.Add(
            "Sales",
            new[] { "Product A", "Product B", "Product C" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Get the collection that controls the appearance of all data labels in the series.
        ChartDataLabelCollection dataLabels = series.DataLabels;

        // Show leader lines so that each label is connected to its slice.
        dataLabels.ShowLeaderLines = true;

        // Position the labels outside the pie slices – this makes the leader lines visible.
        dataLabels.Position = ChartDataLabelPosition.OutsideEnd;

        // OPTIONAL: Increase the visual length of the leader lines by moving the labels
        // farther away from the centre of the pie. This is done by setting an absolute
        // offset for each label (the same offset works for all labels in this simple example).
        for (int i = 0; i < dataLabels.Count; i++)
        {
            ChartDataLabel label = dataLabels[i];
            // Use absolute coordinates measured from the upper‑left corner of the chart.
            label.LeftMode = ChartDataLabelLocationMode.Absolute;
            label.TopMode = ChartDataLabelLocationMode.Absolute;

            // Move each label 20 points to the right and 10 points down.
            // The exact values can be tuned to obtain the desired leader‑line length.
            label.Left += 20;
            label.Top += 10;
        }

        // Save the document to the desired location.
        doc.Save("PieChart_WithLeaderLines.docx");
    }
}
