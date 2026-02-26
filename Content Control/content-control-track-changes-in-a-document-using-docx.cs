using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin tracking revisions. All subsequent edits will be recorded as revisions.
        // The author name "User" will appear in the revision metadata.
        doc.StartTrackRevisions("User");

        // Insert text while tracking is active – this will be stored as an insertion revision.
        builder.Writeln("First line of text.");

        // Insert another paragraph – also recorded as a revision.
        builder.Writeln("Second line, added while tracking.");

        // Stop tracking revisions. Edits after this point will not be recorded.
        doc.StopTrackRevisions();

        // Add content after tracking stopped – this will appear as normal text, not a revision.
        builder.Writeln("This line is not tracked.");

        // Save the document. The file will contain the tracked changes (revisions) made above.
        doc.Save("TrackedChanges.docx");
    }
}
