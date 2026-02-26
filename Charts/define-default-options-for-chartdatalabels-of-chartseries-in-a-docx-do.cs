using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

namespace ChartDataLabelDefaults
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line chart.
            Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
            Chart chart = chartShape.Chart;

            // Remove the demo series that Aspose.Words adds by default.
            chart.Series.Clear();

            // Add a custom series with sample data.
            ChartSeries series = chart.Series.Add("Revenue",
                new[] { "Jan", "Feb", "Mar", "Apr" },
                new[] { 25.6, 21.4, 33.8, 28.1 });

            // Apply default data‑label settings to the series.
            SetDefaultDataLabelOptions(series);

            // If the chart has more than one series, apply the same defaults to each.
            foreach (ChartSeries s in chart.Series)
            {
                SetDefaultDataLabelOptions(s);
            }

            // Save the document.
            doc.Save("ChartWithDefaultDataLabels.docx");
        }

        /// <summary>
        /// Configures a ChartSeries so that its data labels have a consistent set of default options.
        /// </summary>
        /// <param name="series">The series to configure.</param>
        private static void SetDefaultDataLabelOptions(ChartSeries series)
        {
            // Enable data labels for the series.
            series.HasDataLabels = true;

            // Access the collection that represents the series' data labels.
            ChartDataLabelCollection dataLabels = series.DataLabels;

            // Show common elements on every label.
            dataLabels.ShowValue = true;               // Display the numeric value.
            dataLabels.ShowCategoryName = true;        // Display the category (X‑axis) name.
            dataLabels.ShowSeriesName = true;          // Display the series name.
            dataLabels.ShowLeaderLines = true;         // Show leader lines.
            dataLabels.ShowLegendKey = true;           // Show the legend key.
            dataLabels.ShowPercentage = false;         // Do not show percentage (relevant for pie charts).

            // Set a generic number format (two decimal places).
            dataLabels.NumberFormat.FormatCode = "0.00";

            // Use a comma and a space as the separator between label parts.
            dataLabels.Separator = ", ";

            // Set a default font size for readability.
            dataLabels.Font.Size = 10;

            // Optionally, set a default fill and outline for the label callout.
            ChartFormat format = dataLabels.Format;
            format.ShapeType = ChartShapeType.Rectangle; // Simple rectangle label.
            format.Fill.Solid(Color.White);
            format.Stroke.Color = Color.Black;
            format.Stroke.Weight = 0.5; // Stroke width expressed in points.
        }
    }
}
