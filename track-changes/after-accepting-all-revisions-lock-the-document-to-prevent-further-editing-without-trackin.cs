using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial content (this will not be a revision).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe");

        // Make some changes that will be recorded as revisions.
        builder.Writeln("Inserted paragraph while tracking.");

        // Delete the original paragraph to create a deletion revision.
        doc.FirstSection.Body.FirstParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept all revisions so the document contains no pending changes.
        doc.AcceptAllRevisions();

        // Protect the document to make it read‑only. A password is required to unprotect.
        const string password = "SecretPassword";
        doc.Protect(ProtectionType.ReadOnly, password);

        // Save the resulting document.
        doc.Save("LockedDocument.docx");
    }
}
