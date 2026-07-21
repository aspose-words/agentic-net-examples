using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkToPdfExample
{
    public static void Main()
    {
        // Define output paths.
        string outputDir = "Output";
        string pdfPath = Path.Combine(outputDir, "WatermarkedDocument.pdf");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a text watermark.");
        builder.Writeln("The watermark should appear on every page of the PDF.");

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document directly to PDF format.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Simple validation that the PDF file was created.
        if (File.Exists(pdfPath))
        {
            Console.WriteLine($"PDF saved successfully to: {pdfPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the PDF file.");
        }
    }
}
