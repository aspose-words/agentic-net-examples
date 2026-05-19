using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will not be tracked.
        builder.Writeln("Original content.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Perform edits that will be recorded as revisions.
        builder.Writeln("Added line 1.");
        builder.Writeln("Added line 2.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Accept all revisions, removing them from the document.
        doc.AcceptAllRevisions();

        // Ensure all revisions have been accepted.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Protect the document so that any further editing requires tracking changes.
        // Use AllowOnlyRevisions to lock the document for editing without tracking.
        doc.Protect(ProtectionType.AllowOnlyRevisions, "password");

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LockedDocument.docx");
        doc.Save(outputPath);
    }
}
