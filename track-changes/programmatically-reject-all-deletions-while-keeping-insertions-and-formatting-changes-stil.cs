using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will later be deleted.
        builder.Writeln("This paragraph will be deleted.");

        // Start tracking revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Insert new content – this will be recorded as an insertion revision.
        builder.Writeln("This paragraph is an insertion revision.");

        // Delete the first paragraph – this will be recorded as a deletion revision.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[0];
        paragraphToDelete.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only the deletions, leaving insertions and format changes untouched.
        // Iterate backwards to avoid modifying the collection while enumerating.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.RevisionType == RevisionType.Deletion)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");

        // Output the final document text to the console (for verification).
        Console.WriteLine("Final document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
