using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class ChartDataLabelDefaults
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add custom series with sample data.
        ChartSeries series1 = chart.Series.Add("Series 1",
            new[] { "A", "B", "C", "D" },
            new[] { 10.0, 20.0, 30.0, 40.0 });

        ChartSeries series2 = chart.Series.Add("Series 2",
            new[] { "A", "B", "C", "D" },
            new[] { 15.0, 25.0, 35.0, 45.0 });

        // Enable data labels for all series.
        foreach (ChartSeries series in chart.Series)
        {
            series.HasDataLabels = true;
        }

        // Define default font size and color for data labels of every series.
        foreach (ChartSeries series in chart.Series)
        {
            ChartDataLabelCollection dataLabels = series.DataLabels;
            dataLabels.Font.Size = 12;               // Consistent font size.
            dataLabels.Font.Color = Color.DarkBlue;  // Consistent font color.
            dataLabels.ShowValue = true;             // Ensure values are displayed.
        }

        // Save the document.
        doc.Save("ChartWithDefaultDataLabels.docx");
    }
}
