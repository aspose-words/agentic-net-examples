using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // 1. Create a blank document.
        Document blankDoc = new Document();

        // 2. Save the blank document into a memory stream.
        using (MemoryStream createStream = new MemoryStream())
        {
            blankDoc.Save(createStream, SaveFormat.Docx);
            // 3. Reset the stream position to the beginning for loading.
            createStream.Position = 0;

            // 4. Load the document from the memory stream.
            Document doc = new Document(createStream);

            // 5. Start tracking revisions.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // 6. Add a header while tracking is enabled.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Tracked Header");

            // 7. Stop tracking revisions.
            doc.StopTrackRevisions();

            // Optional validation: ensure at least one revision was created.
            if (!doc.HasRevisions || doc.Revisions.Count == 0)
                throw new InvalidOperationException("No revisions were recorded.");

            // 8. Save the final document to disk.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TrackedDocument.docx");
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
