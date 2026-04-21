using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Write("This is a tracked change.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Ensure that a revision was created.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Get the first revision (the insertion we just made).
        Revision revision = doc.Revisions[0];

        // Reject the revision – this removes it from the document.
        revision.Reject();

        // Attempt to accept the same revision again.
        try
        {
            revision.Accept();
            Console.WriteLine("Revision accepted successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected path: the revision has already been rejected.
            Console.WriteLine($"Error while accepting a rejected revision: {ex.Message}");
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
