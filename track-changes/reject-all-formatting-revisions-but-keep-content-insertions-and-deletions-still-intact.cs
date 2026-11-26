using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("This is the original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // ----- Insertion revision -----
        // Insert a new paragraph at the end of the document.
        builder.Writeln("Inserted paragraph.");

        // ----- Deletion revision -----
        // Delete the first paragraph (the original one) while tracking is active.
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        // Use Remove() to correctly generate a deletion revision.
        firstParagraph.Remove();

        // ----- Formatting revision -----
        // Add another paragraph that will receive a formatting change.
        builder.Writeln("Paragraph to format.");
        // Apply a style change (formatting) to the newly added paragraph.
        Paragraph lastParagraph = (Paragraph)doc.GetChild(
            NodeType.Paragraph,
            doc.GetChildNodes(NodeType.Paragraph, true).Count - 1,
            true);
        lastParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1; // Creates a format‑change revision.

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only formatting revisions, keep insertions and deletions.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
