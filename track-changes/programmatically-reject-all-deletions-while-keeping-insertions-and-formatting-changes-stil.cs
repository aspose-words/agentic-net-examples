using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will not be tracked.
        builder.Writeln("Paragraph before tracking.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a new paragraph – this will be an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Apply a formatting change (style change). Formatting changes are not tracked as revisions,
        // but they will remain in the final document.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Delete the original paragraph to create a deletion revision.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[1]; // "Paragraph before tracking."
        paragraphToDelete.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only the deletion revisions, leaving insertions and formatting intact.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.RevisionType == RevisionType.Deletion)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
