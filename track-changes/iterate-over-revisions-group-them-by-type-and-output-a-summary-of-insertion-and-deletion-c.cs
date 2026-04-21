using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text that will NOT be a revision.
        builder.Writeln("This text is not tracked.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert text – this will be recorded as an insertion revision.
        builder.Writeln("First inserted line.");
        builder.Writeln("Second inserted line.");

        // Delete a run (the first paragraph's first run) to create a deletion revision.
        // The run we delete corresponds to the text "This text is not tracked."
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Ensure that revisions were actually created.
        if (!doc.HasRevisions)
        {
            throw new InvalidOperationException("No revisions were generated.");
        }

        // Count revisions by type.
        int insertionCount = 0;
        int deletionCount = 0;

        foreach (Revision rev in doc.Revisions)
        {
            switch (rev.RevisionType)
            {
                case RevisionType.Insertion:
                    insertionCount++;
                    break;
                case RevisionType.Deletion:
                    deletionCount++;
                    break;
                // Other types are ignored for this summary.
            }
        }

        // Output the summary.
        Console.WriteLine($"Revision summary:");
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");

        // Save the document (optional, demonstrates that the document contains revisions).
        doc.Save("RevisionsSummary.docx");
    }
}
