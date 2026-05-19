using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart.
        Chart chart = builder.InsertChart(ChartType.Pie, 500, 300).Chart;

        // Remove the demo series and add custom data.
        chart.Series.Clear();
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels and leader lines.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowLeaderLines = true;
        dataLabels.ShowValue = true;
        dataLabels.ShowPercentage = true;

        // Adjust each label's position to increase the visual length of its leader line.
        // This is done by moving the label outward using absolute coordinates.
        for (int i = 0; i < dataLabels.Count; i++)
        {
            ChartDataLabel label = dataLabels[i];
            label.LeftMode = ChartDataLabelLocationMode.Absolute;
            label.TopMode = ChartDataLabelLocationMode.Absolute;

            // Simple offset calculation – each label is moved a bit farther from the center.
            // The values are arbitrary and can be tuned for better appearance.
            label.Left += 15 * i;   // Move rightward.
            label.Top += 10 * i;    // Move downward.
        }

        // Save the document.
        doc.Save("PieChartLeaderLines.docx");
    }
}
