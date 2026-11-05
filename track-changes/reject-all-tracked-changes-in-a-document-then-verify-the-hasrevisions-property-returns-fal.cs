using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Make some changes while tracking is enabled.
        builder.Writeln("This paragraph is added while tracking.");

        // Delete the first paragraph (will be recorded as a deletion revision).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        firstParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (!doc.HasRevisions)
            throw new Exception("Expected revisions were not created.");

        // Reject all revisions.
        // If Document.RejectAllRevisions() is unavailable, reject each revision individually.
        Revision[] revisions = doc.Revisions.ToArray();
        foreach (Revision rev in revisions)
        {
            rev.Reject();
        }

        // Verify that no revisions remain.
        if (doc.HasRevisions)
            throw new Exception("Revisions were not successfully rejected.");

        // Save the resulting document.
        string outputPath = Path.Combine(Path.GetTempPath(), "TrackChangesReject.docx");
        doc.Save(outputPath);
    }
}
