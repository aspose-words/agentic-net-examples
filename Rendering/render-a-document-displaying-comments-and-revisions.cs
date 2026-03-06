using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions, add a revised paragraph, then stop tracking.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("This paragraph is added as a revision.");
        doc.StopTrackRevisions();

        // Insert a comment into the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
        comment.SetText("This is a comment.");
        builder.CurrentParagraph.AppendChild(comment);

        // Show comments in balloons.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;

        // Configure revision appearance.
        RevisionOptions revOptions = doc.LayoutOptions.RevisionOptions;
        revOptions.ShowInBalloons = ShowInBalloons.Format;          // Show format revisions in balloons.
        revOptions.InsertedTextColor = RevisionColor.BrightGreen;   // Inserted text in bright green.
        revOptions.DeletedTextColor = RevisionColor.Red;           // Deleted text in red.
        revOptions.ShowOriginalRevision = true;                    // Show original text alongside revisions.
        revOptions.ShowRevisionMarks = true;                       // Mark revisions with special formatting.
        revOptions.ShowRevisionBars = true;                        // Show revision bars.

        // Rebuild the layout after changing layout options.
        doc.UpdatePageLayout();

        // Save the document (PDF format demonstrates comment and revision rendering).
        doc.Save("CommentsAndRevisions.pdf");
    }
}
