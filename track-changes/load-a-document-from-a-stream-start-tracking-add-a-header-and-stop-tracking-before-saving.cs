using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add some initial content.
        Document initialDoc = new Document();
        DocumentBuilder initBuilder = new DocumentBuilder(initialDoc);
        initBuilder.Writeln("This is the original content of the document.");

        // Save the document to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            initialDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset the stream before loading.

            // Load the document from the memory stream.
            Document doc = new Document(stream);

            // Start tracking revisions.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // Add a header while tracking is enabled.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Tracked Header");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the resulting document to disk.
            doc.Save("TrackedDocument.docx");
        }
    }
}
