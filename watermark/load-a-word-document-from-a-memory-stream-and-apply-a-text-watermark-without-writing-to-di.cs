using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is a sample document.");

        // Save the document to a memory stream.
        using (MemoryStream sourceStream = new MemoryStream())
        {
            originalDoc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset for reading.

            // Load the document from the memory stream.
            Document doc = new Document(sourceStream);

            // Apply a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document to disk.
            doc.Save("Watermarked.docx");
        }
    }
}
