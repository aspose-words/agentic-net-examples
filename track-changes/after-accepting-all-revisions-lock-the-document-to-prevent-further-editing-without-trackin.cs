using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content (this will not be a revision).
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe");

        // Make changes that will be recorded as revisions.
        builder.Writeln("This paragraph was added while tracking changes.");
        builder.Writeln("Another inserted line.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept all tracked changes so the document has no pending revisions.
        doc.AcceptAllRevisions();

        // Protect the document to prevent further editing without enabling tracking.
        // Using ReadOnly protection; a password is optional.
        doc.Protect(ProtectionType.ReadOnly, "securePassword");

        // Save the resulting document.
        doc.Save("LockedDocument.docx");
    }
}
