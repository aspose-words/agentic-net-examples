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
        builder.Writeln("Original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Make some changes that will be recorded as revisions.
        builder.Writeln("This line is added while tracking is enabled.");
        builder.MoveToDocumentStart();
        builder.Writeln("Inserted at the beginning.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept all revisions.
        doc.AcceptAllRevisions();

        // Verify that revisions have been accepted.
        if (doc.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Revisions were not fully accepted.");
        }

        // Lock the document to prevent further editing without tracking.
        doc.Protect(ProtectionType.ReadOnly, "securePassword");

        // Save the resulting document.
        string outputPath = "TrackedAndLocked.docx";
        doc.Save(outputPath);

        // Output confirmation.
        Console.WriteLine($"Document saved to '{outputPath}'. Protection applied: {doc.ProtectionType != ProtectionType.NoProtection}");
    }
}
