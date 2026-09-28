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
        builder.Writeln("Original text.");

        // Start tracking revisions with a specific author.
        string targetAuthor = "John Doe";
        doc.StartTrackRevisions(targetAuthor, DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Writeln("Added text.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject revisions authored by the target author.
        // Iterate backwards to avoid modifying the collection during enumeration.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.Author == targetAuthor)
                rev.Reject();
        }

        // Validate that the rejected revision is no longer present.
        if (doc.GetText().Contains("Added text"))
            throw new Exception("The revision was not successfully rejected.");

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
