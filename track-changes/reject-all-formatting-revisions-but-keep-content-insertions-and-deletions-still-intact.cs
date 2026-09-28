using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("This is the original paragraph.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // ----- Create an insertion revision -----
        builder.Writeln("This line is inserted while tracking is on.");

        // ----- Create a deletion revision -----
        // Add a paragraph that will be deleted.
        builder.Writeln("This paragraph will be deleted.");
        // Retrieve the paragraph just added.
        Paragraph paraToDelete = doc.LastSection.Body.Paragraphs[doc.LastSection.Body.Paragraphs.Count - 1];
        // Delete it while tracking is active.
        paraToDelete.Remove();

        // ----- Create a formatting revision -----
        // Change the formatting of the first run (make it bold).
        Run firstRun = (Run)doc.GetChild(NodeType.Run, 0, true);
        firstRun.Font.Bold = true; // This generates a FormatChange revision.

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only formatting revisions, keep insertions and deletions.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
