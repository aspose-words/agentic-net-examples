using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Output folder for the generated JPEG images.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "JpgPages");
        Directory.CreateDirectory(outputFolder);

        // Create a simple Word document in memory.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated at runtime.");
        builder.Writeln("It will be saved as high‑quality JPEG images, one per page.");

        // Save the document as PDF first (optional, demonstrates PDF loading).
        string tempPdfPath = Path.Combine(Path.GetTempPath(), "temp_source.pdf");
        doc.Save(tempPdfPath, SaveFormat.Pdf);

        // Load the PDF back into a Document object.
        Document pdfDoc = new Document(tempPdfPath);

        // Configure JPEG save options for high quality.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 100,          // Best quality.
            Resolution = 300,           // High resolution.
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };

        // Save each page of the PDF as a separate JPEG file.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            jpegOptions.PageSet = new PageSet(pageIndex);
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.jpg");
            pdfDoc.Save(outputPath, jpegOptions);
        }

        // Clean up the temporary PDF file.
        if (File.Exists(tempPdfPath))
            File.Delete(tempPdfPath);

        Console.WriteLine($"Saved {pdfDoc.PageCount} JPEG page(s) to \"{outputFolder}\".");
    }
}
