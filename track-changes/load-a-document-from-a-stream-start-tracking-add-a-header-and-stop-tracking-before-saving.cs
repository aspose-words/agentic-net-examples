using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document initialDoc = new Document();
        DocumentBuilder initBuilder = new DocumentBuilder(initialDoc);
        initBuilder.Writeln("Original content of the document.");

        // Save the document to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            initialDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset stream for reading.

            // Load the document from the stream.
            Document doc = new Document(stream);

            // Start tracking revisions.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // Add a header while tracking is enabled.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("This is the header added under tracking.");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the final document to a file.
            doc.Save("TrackedDocument.docx");
        }
    }
}
