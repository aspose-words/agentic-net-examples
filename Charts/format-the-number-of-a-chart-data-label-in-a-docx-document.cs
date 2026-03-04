using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // <-- added
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

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

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with three data points.
        ChartSeries series = chart.Series.Add(
            "Revenue",
            new[] { "January", "February", "March" },
            new[] { 25.611, 21.439, 33.750 });

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Show the value in each data label.
        ChartDataLabelCollection dataLabels = series.DataLabels;
        dataLabels.ShowValue = true;

        // Apply a custom number format to the data labels.
        // This example formats the values as millions of US dollars.
        dataLabels.NumberFormat.FormatCode = "\"US$\" #,##0.000\"M\"";

        // Optionally adjust the font size of the data labels.
        dataLabels.Font.Size = 12;

        // Save the document to a DOCX file.
        doc.Save("ChartDataLabelNumberFormat.docx", SaveFormat.Docx);
    }
}
