using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

class AlignChartLabel
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Clear any demo series.
        chart.Series.Clear();

        // Add a series with categories and values.
        ChartSeries series = chart.Series.Add(
            "Series 1",
            new string[] { "Category 1", "Category 2", "Category 3" },
            new double[] { 4, 5, 6 });

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowValue = true;

        // Align all data labels to the inside base of the column.
        dataLabels.Position = ChartDataLabelPosition.InsideBase;

        // Optionally change the position of the first label.
        dataLabels[0].Position = ChartDataLabelPosition.OutsideEnd;

        // Save the document.
        doc.Save("AlignedChartLabel.docx");
    }
}
