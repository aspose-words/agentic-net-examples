using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample source document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will receive a text watermark.");

        // 2. Save the source document to a memory stream.
        using (MemoryStream sourceStream = new MemoryStream())
        {
            sourceDoc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset for reading.

            // 3. Load the document from the memory stream.
            Document doc = new Document(sourceStream);

            // 4. Apply a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");

            // 5. Save the watermarked document to disk.
            doc.Save("Watermarked.docx");
        }
    }
}
