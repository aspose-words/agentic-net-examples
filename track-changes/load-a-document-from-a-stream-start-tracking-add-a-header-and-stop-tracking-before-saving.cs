using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document sampleDoc = new Document();
        DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
        sampleBuilder.Writeln("Initial content.");

        // Save the sample document to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            sampleDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset stream for reading.

            // Load the document from the stream.
            Document doc = new Document(stream);

            // Start tracking revisions.
            doc.StartTrackRevisions("Demo Author", DateTime.Now);

            // Add a header while tracking is enabled.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("My Header");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the resulting document.
            doc.Save("TrackedDocument.docx");
        }
    }
}
