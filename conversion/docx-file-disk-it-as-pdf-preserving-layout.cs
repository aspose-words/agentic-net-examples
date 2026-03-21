using System;
using System.IO;
using Aspose.Words;

class ConvertDocxToPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This PDF was generated from a DOCX document created in code.");

        // Determine a writable output path (temporary folder).
        string outputPath = Path.Combine(Path.GetTempPath(), "ConvertedDocument.pdf");

        // Ensure the directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // Save the document as PDF. The file extension determines the format.
        doc.Save(outputPath);

        Console.WriteLine($"PDF successfully saved to: {outputPath}");
    }
}
