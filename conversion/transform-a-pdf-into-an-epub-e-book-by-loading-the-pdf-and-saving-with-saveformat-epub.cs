using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder to store temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a simple Word document and save it as PDF.
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);
        builder.Writeln("Sample PDF content for EPUB conversion.");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Step 2: Load the generated PDF.
        Document pdfDoc = new Document(pdfPath);

        // Step 3: Convert the PDF to EPUB.
        string epubPath = Path.Combine(outputDir, "sample.epub");
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Step 4: Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
        {
            throw new InvalidOperationException("EPUB conversion failed: output file not found.");
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("PDF successfully converted to EPUB at: " + epubPath);
    }
}
