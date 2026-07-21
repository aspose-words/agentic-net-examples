using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path where the resulting document will be saved.
        const string outputPath = "LockedDocument.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (this will not be a revision).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe");

        // Make some changes while tracking is active.
        builder.Writeln("Added paragraph while tracking.");

        // Delete the original paragraph to generate a deletion revision.
        Paragraph originalParagraph = doc.FirstSection.Body.Paragraphs[0];
        originalParagraph.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Accept all revisions so the document becomes clean.
        doc.AcceptAllRevisions();

        // Protect the document so that further editing is only possible when track changes is enabled.
        // ProtectionType.AllowOnlyRevisions permits only revision marks.
        doc.Protect(ProtectionType.AllowOnlyRevisions, "securePassword");

        // Save the locked document.
        doc.Save(outputPath);
    }
}
