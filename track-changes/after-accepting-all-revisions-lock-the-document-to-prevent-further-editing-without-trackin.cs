using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make some changes while tracking is enabled.
        builder.Writeln("This paragraph is added while tracking.");
        // Delete the original paragraph to generate a deletion revision.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept all revisions that were generated.
        doc.AcceptAllRevisions();

        // Protect the document to prevent further editing without tracking.
        doc.Protect(ProtectionType.ReadOnly, "password");

        // Save the resulting document.
        doc.Save("LockedDocument.docx");
    }
}
