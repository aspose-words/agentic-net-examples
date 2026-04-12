using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToImagesExample
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF document.");
        builder.Writeln("It will be converted to JPG and PNG images.");
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document for conversion.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Convert the PDF to a JPEG image.
        // -----------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: set resolution or quality if desired.
            Resolution = 300,
            JpegQuality = 90
        };
        string jpegPath = Path.Combine(artifactsDir, "Sample.jpg");
        pdfDocument.Save(jpegPath, jpegOptions);

        // Verify JPEG output.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("JPEG image was not created.");

        // -----------------------------------------------------------------
        // 4. Convert the PDF to a PNG image.
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set resolution.
            Resolution = 300
        };
        string pngPath = Path.Combine(artifactsDir, "Sample.png");
        pdfDocument.Save(pngPath, pngOptions);

        // Verify PNG output.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("PNG image was not created.");

        // -----------------------------------------------------------------
        // 5. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine("PDF successfully converted to JPG and PNG.");
        Console.WriteLine($"JPG file: {jpegPath}");
        Console.WriteLine($"PNG file: {pngPath}");
    }
}
