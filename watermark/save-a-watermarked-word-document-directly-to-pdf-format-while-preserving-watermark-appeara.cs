using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkToPdfExample
{
    public static void Main()
    {
        // Define output PDF file path.
        string outputPath = "WatermarkedDocument.pdf";

        // Create a new blank Word document.
        Document doc = new Document();

        // Add some sample content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a watermark.");
        builder.Writeln("The watermark should appear in the saved PDF.");

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document directly to PDF format, preserving the watermark.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Optional: verify that the PDF file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
    }
}
