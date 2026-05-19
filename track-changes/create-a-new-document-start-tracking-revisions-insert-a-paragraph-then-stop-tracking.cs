using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions with a specified author and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a paragraph while tracking is enabled – this will be recorded as a revision.
        builder.Writeln("This paragraph is inserted while tracking changes.");

        // Stop tracking revisions so subsequent edits are not recorded.
        doc.StopTrackRevisions();

        // Save the document to verify that the revision was created.
        doc.Save("TrackedDocument.docx");
    }
}
