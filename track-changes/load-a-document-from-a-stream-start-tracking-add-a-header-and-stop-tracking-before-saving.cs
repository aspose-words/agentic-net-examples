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
        using (MemoryStream stream = new MemoryStream())
        {
            blankDoc.Save(stream, SaveFormat.Docx);
            // Reset the stream position before loading.
            stream.Position = 0;

            // 3. Load the document from the memory stream.
            Document doc = new Document(stream);

            // 4. Start tracking revisions.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // 5. Add a header while tracking is enabled.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Tracked Header");

            // 6. Stop tracking revisions.
            doc.StopTrackRevisions();

            // 7. Save the resulting document to disk.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TrackedDocument.docx");
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
