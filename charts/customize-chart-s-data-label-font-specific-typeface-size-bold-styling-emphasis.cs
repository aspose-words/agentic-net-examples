using System;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class CustomizeChartDataLabelFont
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove any demo series that Aspose adds by default.
        chart.Series.Clear();

        // Add a custom series with three categories.
        ChartSeries series = chart.Series.Add(
            "Sample Series",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 10.0, 20.0, 30.0 });

        // Enable data labels for the series.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowValue = true; // Show the numeric values.

        // Customize the font of all data labels in this series.
        dataLabels.Font.Name = "Calibri";   // Typeface.
        dataLabels.Font.Size = 14;          // Font size.
        dataLabels.Font.Bold = true;        // Bold styling.

        // Save the document.
        doc.Save("CustomizedDataLabelFont.docx");
    }
}
