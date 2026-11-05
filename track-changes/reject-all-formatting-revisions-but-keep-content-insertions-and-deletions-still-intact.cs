using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph that will later be deleted.
        builder.Writeln("This paragraph will be deleted.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert a new paragraph – this creates an insertion revision.
        builder.Writeln("This is an inserted paragraph.");

        // Apply a formatting change to the inserted paragraph – creates a format change revision.
        // The builder is positioned at the end of the inserted paragraph, so the formatting change
        // will be recorded as a revision.
        builder.Font.Bold = true;

        // Delete the first paragraph – creates a deletion revision.
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        firstParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only formatting revisions, keep insertions and deletions intact.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
            {
                rev.Reject();
            }
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
