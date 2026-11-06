using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Start tracking revisions.
        string author = "Demo Author";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Insert a paragraph while tracking is enabled.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a newly inserted paragraph while tracking changes.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document.
        string outputPath = "TrackedDocument.docx";
        doc.Save(outputPath);

        // Verify that a revision was recorded.
        if (doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were recorded.");
        }
    }
}
