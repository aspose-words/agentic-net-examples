using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document.");

        // Save the document to a MemoryStream.
        using (MemoryStream stream = new MemoryStream())
        {
            doc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset stream for reading.

            // Load the document from the MemoryStream.
            Document loadedDoc = new Document(stream);

            // Apply a text watermark.
            loadedDoc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document to disk.
            loadedDoc.Save("Watermarked.docx");
        }
    }
}
