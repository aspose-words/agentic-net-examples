using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (this will be the text we later delete).
        builder.Writeln("Original paragraph.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // ----- Insertion revision -----
        builder.Writeln("Inserted paragraph.");

        // ----- Deletion revision -----
        // Delete the original paragraph while tracking is active.
        Paragraph originalParagraph = doc.FirstSection.Body.Paragraphs[0];
        originalParagraph.Remove();

        // ----- Formatting change revision -----
        builder.Font.Bold = true;
        builder.Writeln("Bold formatted paragraph.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only deletions, keep insertions and formatting changes.
        // Create a snapshot of the revisions to avoid modifying the collection during enumeration.
        Revision[] revisionsSnapshot = doc.Revisions.Cast<Revision>().ToArray();
        foreach (Revision rev in revisionsSnapshot)
        {
            if (rev.RevisionType == RevisionType.Deletion)
                rev.Reject();
        }

        // Verify that no deletion revisions remain.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Deletion)
                throw new Exception("A deletion revision was not rejected properly.");
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
