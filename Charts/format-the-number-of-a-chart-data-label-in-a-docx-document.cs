using System;
using Aspose.Words;
using Aspose.Words.Drawing; // Added for Shape
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the default demo series.
        chart.Series.Clear();

        // Add a custom series with sample data.
        ChartSeries series = chart.Series.Add(
            "Revenue",
            new[] { "January", "February", "March" },
            new[] { 25.611, 21.439, 33.750 });

        // Enable data labels for the series and show the values.
        series.HasDataLabels = true;
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowValue = true;

        // Set a custom number format for the data labels.
        // This example formats the values as US dollars with two decimal places.
        dataLabels.NumberFormat.FormatCode = "\"US$\" #,##0.00";

        // Optional: adjust the font size of the data labels.
        dataLabels.Font.Size = 10;

        // Save the document to a DOCX file.
        doc.Save("ChartDataLabelNumberFormat.docx");
    }
}
