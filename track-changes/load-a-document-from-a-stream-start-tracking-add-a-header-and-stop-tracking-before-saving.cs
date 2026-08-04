using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // 1. Create a blank document and add a simple paragraph.
        Document initialDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(initialDoc);
        builder.Writeln("Initial content before tracking.");

        // 2. Save the document into a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            initialDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset for reading.

            // 3. Load the document from the stream.
            Document doc = new Document(stream);

            // 4. Start tracking revisions.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // 5. Add a header while tracking is enabled.
            DocumentBuilder headerBuilder = new DocumentBuilder(doc);
            headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            headerBuilder.Writeln("Tracked Header Text");

            // 6. Stop tracking revisions.
            doc.StopTrackRevisions();

            // 7. Save the resulting document to disk.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TrackedChanges.docx");
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
