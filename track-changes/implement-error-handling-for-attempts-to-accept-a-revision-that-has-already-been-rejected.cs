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
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make a change that will generate a revision (insertion).
        builder.Writeln("Inserted paragraph while tracking.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Ensure there is at least one revision.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Get the first revision (the insertion we just made).
        Revision revision = doc.Revisions[0];

        // Reject the revision.
        revision.Reject();

        // Attempt to accept the same revision after it has been rejected.
        try
        {
            revision.Accept();
        }
        catch (InvalidOperationException ex)
        {
            // Expected error handling for already rejected revision.
            Console.WriteLine("Error: Cannot accept a revision that has already been rejected.");
            Console.WriteLine("Exception message: " + ex.Message);
        }

        // Save the resulting document.
        doc.Save("TrackChangesResult.docx");
    }
}
