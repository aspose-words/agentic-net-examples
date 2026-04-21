using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial text – this will not be a revision.
        builder.Writeln("This is the original paragraph.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Perform edits that will be recorded as revisions.
        builder.Writeln("First tracked change.");
        builder.Writeln("Second tracked change.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (!doc.HasRevisions)
            throw new Exception("Expected revisions were not generated.");

        // Accept all tracked changes, removing the revision collection.
        doc.AcceptAllRevisions();

        // Ensure all revisions have been cleared.
        if (doc.HasRevisions)
            throw new Exception("Revisions were not fully accepted.");

        // Protect the document to block editing without tracking.
        // Using ReadOnly protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "myPassword");

        // Save the resulting document.
        doc.Save("LockedDocument.docx");
    }
}
