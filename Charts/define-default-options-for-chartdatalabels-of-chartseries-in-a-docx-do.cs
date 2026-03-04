using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Added for Shape
using Aspose.Words.Drawing.Charts;

class ChartDataLabelsDefault
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart (column type) into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 400);
        Chart chart = chartShape.Chart;

        // Iterate through all series in the chart and set default data label options.
        foreach (ChartSeries series in chart.Series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Access the collection that controls the default settings for all data labels in the series.
            ChartDataLabelCollection dataLabels = series.DataLabels;

            // Set default visibility options.
            dataLabels.ShowValue = true;            // Show the numeric value.
            dataLabels.ShowSeriesName = true;       // Show the series name.
            dataLabels.ShowCategoryName = true;     // Show the category name.

            // Set other default options as needed.
            dataLabels.ShowLeaderLines = false;     // No leader lines.
            dataLabels.ShowLegendKey = false;       // No legend key.
            dataLabels.ShowPercentage = false;      // No percentage (relevant for pie charts).

            // Define the default separator string used when multiple parts are displayed.
            dataLabels.Separator = ", ";

            // Example of setting the default position (optional, depends on chart type).
            // dataLabels.Position = ChartDataLabelPosition.Center; // Uncomment if the enum is available.
        }

        // Save the document.
        doc.Save("ChartDataLabelsDefault.docx");
    }
}
