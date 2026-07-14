using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.pdf");
        string pngPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");

        // -----------------------------------------------------------------
        // Step 1: Create a Word document with vector graphics (a chart).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a title.
        builder.Writeln("Vector Graphics Sample");

        // Insert a pie chart – charts are vector graphics.
        Chart chart = builder.InsertChart(ChartType.Pie, 500, 400).Chart;
        chart.Title.Text = "Sample Chart";
        chart.Series.Clear();

        // Add a series with sample data.
        ChartSeries series = chart.Series.Add("Series 1",
            new string[] { "A", "B", "C" },
            new double[] { 30, 45, 25 });

        // (Optional) Enable data labels if needed.
        // The ChartSeries class does not expose a HasDataLabel property in this version,
        // so this line is omitted.

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert the first page to a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render at 300 DPI for high resolution.
            Resolution = 300,
            // Use high‑quality rendering algorithms.
            UseHighQualityRendering = true,
            // Render only the first page (zero‑based index).
            PageSet = new PageSet(0)
        };

        pdfDoc.Save(pngPath, pngOptions);

        // Verify that the PNG was created and contains data.
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("The PNG file was not created or is empty.");

        // Indicate successful completion.
        Console.WriteLine("PDF successfully converted to high‑resolution PNG.");
    }
}
