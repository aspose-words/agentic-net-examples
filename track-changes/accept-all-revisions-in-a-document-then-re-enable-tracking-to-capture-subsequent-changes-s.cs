using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content without tracking (no revisions will be created).
        builder.Writeln("Initial paragraph. ");

        // Start tracking revisions with the first author.
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Make changes while tracking – these will be recorded as revisions.
        builder.Writeln("First revision paragraph. ");
        builder.Writeln("Another revision paragraph. ");

        // Stop tracking – further changes will not be recorded as revisions.
        doc.StopTrackRevisions();

        // At this point the document has revisions.
        int revisionCountBefore = doc.Revisions.Count;
        if (revisionCountBefore == 0)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Accept all existing revisions, removing them from the document.
        doc.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Start tracking again with a different author to capture new changes separately.
        doc.StartTrackRevisions("Author2", DateTime.Now);

        // Add more content – these will form a new set of revisions.
        builder.Writeln("Second tracking paragraph. ");
        builder.Writeln("Yet another paragraph after re‑enabling tracking. ");

        // Stop tracking once more.
        doc.StopTrackRevisions();

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "AcceptAndTrackRevisions.docx");
        doc.Save(outputPath);
    }
}
