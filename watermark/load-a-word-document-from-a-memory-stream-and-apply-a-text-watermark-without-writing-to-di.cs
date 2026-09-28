using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document.");

        // Save the sample document to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            sampleDoc.Save(stream, SaveFormat.Docx);
            // Reset the stream position before loading.
            stream.Position = 0;

            // Load the document from the memory stream.
            Document loadedDoc = new Document(stream);

            // Apply a text watermark.
            loadedDoc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document to disk.
            loadedDoc.Save("Watermarked.docx");
        }
    }
}
