using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial paragraph.
        builder.Writeln("This is the original paragraph.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Make a change while tracking is enabled – this will create a revision.
        builder.Writeln("This line is added while tracking is ON.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Capture the number of revisions after the first change.
        int revisionsAfterFirstEdit = doc.Revisions.Count;
        Console.WriteLine($"Revisions after first edit (tracking ON): {revisionsAfterFirstEdit}");

        // Make another change after tracking has been stopped.
        builder.Writeln("This line is added after tracking is OFF.");

        // Verify that no new revisions were recorded.
        int revisionsAfterSecondEdit = doc.Revisions.Count;
        Console.WriteLine($"Revisions after second edit (tracking OFF): {revisionsAfterSecondEdit}");

        if (revisionsAfterSecondEdit != revisionsAfterFirstEdit)
            throw new InvalidOperationException("New revisions were recorded after stopping tracking.");

        // Save the document to verify the result manually if needed.
        doc.Save("TrackChangesDemo.docx");
    }
}
