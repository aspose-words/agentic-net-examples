using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;

public class EnableLeaderLinesPieChart
{
    public static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series and add a custom one.
        chart.Series.Clear();
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 30.0, 45.0, 25.0 });

        // Enable data labels and show leader lines.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowLeaderLines = true;
        dataLabels.ShowValue = true;
        dataLabels.ShowPercentage = true;

        // Adjust label positions to increase leader line length.
        // Offset each label outward by a fixed amount.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            ChartDataLabel label = dataLabels[i];
            label.Left += 15;                     // Move label to the right.
            label.Top += 15;                      // Move label downward.
            label.LeftMode = ChartDataLabelLocationMode.Offset;
            label.TopMode = ChartDataLabelLocationMode.Offset;
        }

        // Save the document.
        doc.Save("EnableLeaderLinesPieChart.docx");
    }
}
