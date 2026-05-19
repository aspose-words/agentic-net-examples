using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will later be partially deleted.
        builder.Write("This is the original paragraph. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Insert new text – this will be an insertion revision.
        builder.Write("Inserted sentence. ");

        // Delete a portion of the original text – this will be a deletion revision.
        // Remove the first run (the original sentence) to create a deletion revision.
        Node firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Reject only the deletion revisions, leaving insertions and any formatting changes untouched.
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
