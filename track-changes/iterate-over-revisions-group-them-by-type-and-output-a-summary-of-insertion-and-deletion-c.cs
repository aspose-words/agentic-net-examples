using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text that will later be deleted.
        builder.Write("This text will be deleted. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Write("This is an inserted sentence. ");

        // Delete the first run (the text written before tracking started) – this creates a deletion revision.
        Run firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates that revisions are persisted).
        doc.Save("RevisionsSample.docx");

        // Iterate over all revisions and count insertions and deletions.
        int insertionCount = 0;
        int deletionCount = 0;

        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion)
                insertionCount++;
            else if (rev.RevisionType == RevisionType.Deletion)
                deletionCount++;
        }

        // Output the summary.
        Console.WriteLine($"Total insertion revisions: {insertionCount}");
        Console.WriteLine($"Total deletion revisions: {deletionCount}");
    }
}
