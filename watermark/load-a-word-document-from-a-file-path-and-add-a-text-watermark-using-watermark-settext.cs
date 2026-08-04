using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the source and output documents
        string sourcePath = "Sample.docx";
        string outputPath = "Watermarked.docx";

        // Create a sample document if it does not already exist
        if (!File.Exists(sourcePath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document.");
            sampleDoc.Save(sourcePath);
        }

        // Load the existing document
        Document doc = new Document(sourcePath);

        // Add a text watermark
        doc.Watermark.SetText("Confidential");

        // Save the document with the watermark applied
        doc.Save(outputPath);

        // Indicate that the process has completed
        Console.WriteLine($"Watermarked document saved to '{outputPath}'.");
    }
}
