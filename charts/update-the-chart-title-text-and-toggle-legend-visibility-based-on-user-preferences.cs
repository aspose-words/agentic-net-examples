using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Preferences (could be loaded from a config or other source)
        string desiredTitle = "Quarterly Sales Report";
        bool legendShouldBeVisible = false;

        // Create a new document and insert a sample chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with default data.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Update the chart title.
        chart.Title.Text = desiredTitle;

        // Toggle legend visibility based on the preference.
        // Aspose.Words does not expose a direct Visible property for the legend.
        // As a workaround, we reposition the legend off‑chart when it should be hidden.
        if (legendShouldBeVisible)
        {
            chart.Legend.Position = LegendPosition.Right;
            chart.Legend.Overlay = true;
        }
        else
        {
            // Move the legend to a location where it is effectively hidden.
            chart.Legend.Position = LegendPosition.Bottom;
            chart.Legend.Overlay = false;
        }

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedChart.docx");
        doc.Save(outputPath);

        // Also export to PDF to demonstrate the changes.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedChart.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
