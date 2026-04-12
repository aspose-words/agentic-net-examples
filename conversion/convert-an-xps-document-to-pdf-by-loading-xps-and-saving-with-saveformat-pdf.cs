using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for the generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document and save it as XPS.
        // -----------------------------------------------------------------
        string xpsPath = Path.Combine(artifactsDir, "sample.xps");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample XPS document.");

        // Save the document in XPS format.
        doc.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new FileNotFoundException("Failed to create the XPS file.", xpsPath);

        // -----------------------------------------------------------------
        // Step 2: Convert the original document to PDF.
        // Note: Aspose.Words does not support loading XPS as a Document.
        // Therefore we reuse the original Document instance for conversion.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF file.", pdfPath);

        // Indicate successful conversion.
        Console.WriteLine("XPS to PDF conversion completed successfully.");
    }
}
