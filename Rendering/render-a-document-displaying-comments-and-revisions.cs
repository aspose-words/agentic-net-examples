using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

class RenderCommentsAndRevisions
{
    static void Main()
    {
        // Define output folder.
        string artifactsDir = @"C:\Output\";
        Directory.CreateDirectory(artifactsDir);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("This is the original paragraph.");

        // Insert a comment.
        Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
        comment.SetText("This is a comment.");
        builder.CurrentParagraph.AppendChild(comment);

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        // Make a revision: insert new text.
        builder.Writeln("This line was added as a revision.");
        // End tracking revisions.
        doc.StopTrackRevisions();

        // Configure how comments are displayed (balloons in the margin).
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;

        // Configure revision appearance.
        RevisionOptions revOptions = doc.LayoutOptions.RevisionOptions;
        revOptions.ShowOriginalRevision = true;               // Show original text alongside revised.
        revOptions.ShowRevisionBars = true;                  // Show revision bars.
        revOptions.ShowInBalloons = ShowInBalloons.Format;   // Show format revisions in balloons.
        revOptions.CommentColor = RevisionColor.BrightGreen; // Color for comment balloons.

        // Apply layout changes.
        doc.UpdatePageLayout();

        // Save the document to PDF (comments and revisions will be rendered).
        doc.Save(Path.Combine(artifactsDir, "CommentsAndRevisions.pdf"));
    }
}
