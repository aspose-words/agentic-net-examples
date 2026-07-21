using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content (no revision yet).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // ----- Insertion revision -----
        builder.Writeln("Inserted paragraph.");

        // ----- Deletion revision -----
        // Remove the first paragraph that was added before tracking started.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();

        // ----- Formatting change (not recorded as a revision) -----
        builder.Font.Bold = true;
        builder.Writeln("Bold text.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only formatting revisions, leaving insertions and deletions untouched.
        // Iterate backwards to safely modify the collection while iterating.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
