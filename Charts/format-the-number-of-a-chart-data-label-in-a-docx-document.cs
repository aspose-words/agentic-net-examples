using System;
using Aspose.Words;
using Aspose.Words.Drawing; // <-- added
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

class ChartDataLabelNumberFormatExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Add a custom series with categories and numeric values.
        chart.Series.Add("Revenue",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new double[] { 125000.75, 98000.5, 143200.25, 167500.0 });

        // Enable data labels for the series.
        ChartSeries series = chart.Series[0];
        series.HasDataLabels = true;

        // Show the value in each data label.
        series.DataLabels.ShowValue = true;

        // Apply a custom number format to all data labels of the series.
        // Example format: US dollars with two decimal places.
        series.DataLabels.NumberFormat.FormatCode = "\"US$\" #,##0.00";

        // Optionally, adjust the font size of the data labels.
        series.DataLabels.Font.Size = 10;

        // Save the document to a DOCX file.
        doc.Save("ChartDataLabelNumberFormat.docx", SaveFormat.Docx);
    }
}
