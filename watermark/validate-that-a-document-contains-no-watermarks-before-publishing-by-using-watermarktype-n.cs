using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path of the sample document.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");

        // Create a blank document and add a simple paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a test document without a watermark.");

        // Save the document.
        doc.Save(docPath);

        // Load the document back (simulating a publishing workflow).
        Document loadedDoc = new Document(docPath);

        // Validate that the document has no watermark.
        if (loadedDoc.Watermark.Type == WatermarkType.None)
        {
            Console.WriteLine("Validation passed: No watermark is present.");
        }
        else
        {
            Console.WriteLine("Validation failed: Watermark detected.");
        }
    }
}
