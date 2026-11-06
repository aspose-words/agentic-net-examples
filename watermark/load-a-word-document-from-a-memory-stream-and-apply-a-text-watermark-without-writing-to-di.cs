using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document.");

        // Save the document to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            sampleDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset stream position for reading.

            // Load the document from the memory stream.
            Document loadedDoc = new Document(stream);

            // Apply a text watermark.
            loadedDoc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document to disk.
            string outputPath = "Watermarked.docx";
            loadedDoc.Save(outputPath);

            // Simple validation that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Watermark applied successfully. Output saved to '{outputPath}'.");
            }
        }
    }
}
