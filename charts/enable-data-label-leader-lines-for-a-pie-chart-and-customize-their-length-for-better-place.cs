using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;

public class LeaderLinesPieChart
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart.
        Chart chart = builder.InsertChart(ChartType.Pie, 500, 300).Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;

        // Show leader lines, values and percentages.
        dataLabels.ShowLeaderLines = true;
        dataLabels.ShowValue = true;
        dataLabels.ShowPercentage = true;

        // Customize the position of each data label to increase leader line length.
        // Use absolute positioning for precise control.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            ChartDataLabel label = dataLabels[i];
            // Move each label outward by a fixed offset.
            // The offset values are arbitrary and can be tuned as needed.
            label.Left += 20 * i;
            label.Top += 20 * i;
            label.LeftMode = ChartDataLabelLocationMode.Absolute;
            label.TopMode = ChartDataLabelLocationMode.Absolute;
        }

        // Save the document.
        doc.Save("LeaderLinesPieChart.docx");
    }
}
