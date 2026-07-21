using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the input and output documents.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleInput.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleOutput.docx");

        // Create a simple source document if it does not already exist.
        if (!File.Exists(inputPath))
        {
            Document sourceDoc = new Document();
            var builder = new DocumentBuilder(sourceDoc);
            builder.Writeln("This is a sample document.");
            sourceDoc.Save(inputPath);
        }

        // Load the existing document from the file system.
        Document doc = new Document(inputPath);

        // Add a text watermark to the loaded document.
        doc.Watermark.SetText("Confidential");

        // Save the document with the watermark applied.
        doc.Save(outputPath);
    }
}
