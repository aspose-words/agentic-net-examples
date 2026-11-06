using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial content with separate runs for "Hello " and "world."
        builder.Write("Hello ");
        builder.Write("world.");

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // ----- Deletion: remove the run that contains "world."
        // This creates a deletion revision.
        Run worldRun = doc.FirstSection.Body.FirstParagraph.Runs[1];
        worldRun.Remove();

        // ----- Insertion: add new text.
        // This creates an insertion revision.
        builder.MoveTo(doc.FirstSection.Body.FirstParagraph);
        builder.Write("Inserted text ");

        // ----- Formatting change: make "Hello " bold.
        // This creates a format change revision.
        Run helloRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        helloRun.Font.Bold = true;

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only deletions, keep insertions and formatting changes.
        // Iterate backwards to avoid modifying the collection during enumeration.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.RevisionType == RevisionType.Deletion)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
