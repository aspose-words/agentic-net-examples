using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document created for watermark demonstration.");

        // Save the sample document to a memory stream.
        using (MemoryStream sourceStream = new MemoryStream())
        {
            sampleDoc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset stream before reading.

            // Load the document from the memory stream.
            Document loadedDoc = new Document(sourceStream);

            // Apply a text watermark.
            loadedDoc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document to disk.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Watermarked.docx");
            loadedDoc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
