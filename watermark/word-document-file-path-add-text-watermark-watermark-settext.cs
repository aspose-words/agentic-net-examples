using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph so the document has some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Add a plain text watermark to every page of the document.
        // The watermark text must be between 1 and 200 characters.
        doc.Watermark.SetText("Confidential");

        // Determine output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample_With_Watermark.docx");

        // Save the modified document.
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
