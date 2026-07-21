using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document blankDoc = new Document();

        // Save the blank document into a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            blankDoc.Save(stream, SaveFormat.Docx);
            // Reset the stream position before loading.
            stream.Position = 0;

            // Load the document from the memory stream.
            Document doc = new Document(stream);

            // Start tracking revisions with a specific author.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // Add a header while tracking is enabled.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Tracked Header");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the resulting document to a file.
            doc.Save("TrackedDocument.docx");
        }
    }
}
