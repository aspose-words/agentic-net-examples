using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "sample.docx";
        const string outputPath = "result.pdf";

        // -----------------------------------------------------------------
        // Create a sample DOCX file containing a placeholder.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear _CustomerName_,");
        builder.Writeln("Thank you for your purchase.");
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the DOCX, replace the placeholder, and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);
        int replacements = doc.Range.Replace("_CustomerName_", "John Doe");
        if (replacements == 0)
        {
            throw new InvalidOperationException("Placeholder was not found in the document.");
        }

        doc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }

        // Optional: indicate success (no console interaction required).
        Console.WriteLine("Conversion completed successfully.");
    }
}
