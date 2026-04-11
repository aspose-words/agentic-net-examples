using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.pdf");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Converted.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a simple PDF document to act as the source file.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF created with Aspose.Words.");
        builder.Writeln("It will be converted to DOCX format.");

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException($"Failed to convert PDF to DOCX at '{docxPath}'.");

        // Optional: Output the result paths (no interactive prompts).
        Console.WriteLine($"PDF created at: {pdfPath}");
        Console.WriteLine($"DOCX converted at: {docxPath}");
    }
}
