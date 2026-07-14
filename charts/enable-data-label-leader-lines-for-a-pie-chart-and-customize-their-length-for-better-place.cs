using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a pie chart.
        Shape chartShape = builder.InsertChart(ChartType.Pie, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add custom data.
        string[] categories = { "Apples", "Bananas", "Cherries" };
        double[] values = { 30, 45, 25 };
        ChartSeries series = chart.Series.Add("Fruits", categories, values);

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowLeaderLines = true;   // Show leader lines.
        dataLabels.ShowValue = true;         // Show the numeric value.
        dataLabels.ShowPercentage = true;    // Show the percentage.
        dataLabels.Separator = "; ";         // Separator between value and percentage.

        // Increase the distance between the data labels and the pie slices.
        // This is achieved by setting the series' Explosion property, which moves
        // each slice (and its leader line) outward from the centre.
        series.Explosion = 20; // Adjust this value as needed for longer leader lines.

        // Save the document.
        doc.Save("DataLabelsLeaderLinesPieChart.docx");
    }
}
