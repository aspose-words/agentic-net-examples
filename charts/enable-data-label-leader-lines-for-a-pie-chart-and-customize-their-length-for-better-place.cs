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

        // Insert a pie chart.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series.
        chart.Series.Clear();

        // Add a custom series with categories and values.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Apples", "Bananas", "Cherries" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels and show leader lines.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowLeaderLines = true;
        dataLabels.ShowValue = true;
        dataLabels.ShowPercentage = true;

        // Customize the length of the leader lines by moving each label outward.
        // This is done by setting absolute positions for the labels.
        for (int i = 0; i < dataLabels.Count; i++)
        {
            ChartDataLabel label = dataLabels[i];
            // Position the label farther from the center based on its index.
            double offset = 20.0 * (i + 1);
            label.Left = offset;
            label.Top = offset;
            label.LeftMode = ChartDataLabelLocationMode.Absolute;
            label.TopMode = ChartDataLabelLocationMode.Absolute;
        }

        // Save the document.
        doc.Save("EnableLeaderLinesPieChart.docx");
    }
}
