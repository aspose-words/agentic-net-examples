using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Writeln("Inserted paragraph.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Ensure a revision was created.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Get the first revision.
        Revision revision = doc.Revisions[0];

        // Reject the revision.
        revision.Reject();

        // Attempt to accept the same revision after it has been rejected.
        try
        {
            revision.Accept();
            Console.WriteLine("Revision accepted successfully (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accepting already rejected revision: {ex.GetType().Name} - {ex.Message}");
        }

        // Verify that the revision collection is now empty.
        Console.WriteLine($"Revisions count after reject: {doc.Revisions.Count}");
    }
}
