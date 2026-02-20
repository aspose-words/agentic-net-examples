using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

namespace AsposeWordsChartDataLabelExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart into the document.
            // Width = 432 points, Height = 252 points.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = chartShape.Chart;

            // Remove the default series that Aspose.Words creates.
            chart.Series.Clear();

            // Add a single series with three data points.
            string[] categories = new string[] { "Category A", "Category B", "Category C" };
            double[] values = new double[] { 10, 20, 30 };
            ChartSeries series = chart.Series.Add("My Series", categories, values);

            // Enable data labels for the whole series.
            series.HasDataLabels = true;

            // Access the collection of data labels for the series.
            ChartDataLabelCollection dataLabels = series.DataLabels;

            // Show the value for each data label.
            dataLabels.ShowValue = true;

            // Change the position of all data labels to be inside the base of the column.
            dataLabels.Position = ChartDataLabelPosition.InsideBase;

            // Work with a single data label (e.g., the second point, index 1).
            ChartDataLabel label = dataLabels[1];

            // Set additional properties for this label.
            label.ShowCategoryName = true;   // Show the category name.
            label.ShowSeriesName = true;     // Show the series name.
            label.ShowLeaderLines = true;    // Show leader lines.
            label.Separator = " | ";         // Use a custom separator.

            // Change the font color of this label to dark red.
            label.Font.Color = Color.DarkRed;

            // Change the fill color of the label background.
            label.Format.Fill.Color = Color.LightYellow;

            // Optionally, hide the label completely.
            // label.IsHidden = true;

            // Save the document to disk.
            doc.Save("ChartDataLabelExample.docx");
        }
    }
}
