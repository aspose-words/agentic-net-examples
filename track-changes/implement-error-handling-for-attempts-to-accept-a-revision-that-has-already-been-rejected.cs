using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello");

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make a change that will generate a revision.
        builder.Writeln("This line is added while tracking.");

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Verify that a revision exists.
        if (doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision, but none were found.");
        }

        // Get the first revision.
        Revision revision = doc.Revisions[0];

        // Reject the revision.
        revision.Reject();

        // Attempt to accept the same revision again and handle the expected error.
        try
        {
            revision.Accept();
            Console.WriteLine("Revision accepted (unexpected).");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine("Error: Cannot accept a revision that has already been rejected. " + ex.Message);
        }

        // Save the final document (optional).
        doc.Save("Output.docx");
    }
}
