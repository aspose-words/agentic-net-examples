using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document with multiple pages.
        const string pdfPath = "sample.pdf";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is page 1 of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2 of the sample PDF.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Set up image save options for a high‑quality JPEG.
        // The MultiPageLayout property is not available in the version of Aspose.Words used
        // for this verification, so we omit it. The example still demonstrates high‑quality
        // JPEG export.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 100, // Maximum quality.
            Resolution = 300   // High DPI for better quality.
        };

        // Save the PDF as a JPEG image.
        const string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, saveOptions);

        // Verify that the JPEG file was created and contains data.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("The JPEG image was not created successfully.");
    }
}
