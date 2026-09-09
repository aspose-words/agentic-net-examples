using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample DOCX document.");
        builder.Writeln("It will be converted to PDF while preserving layout.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX file from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Save the document as PDF.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 4. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }

        // Optional: Inform the user (no interactive wait required).
        Console.WriteLine($"Conversion completed successfully. PDF saved to: {outputPath}");
    }
}
