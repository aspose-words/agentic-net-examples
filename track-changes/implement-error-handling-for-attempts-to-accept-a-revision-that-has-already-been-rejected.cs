using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Make a change that will generate a revision (insertion).
        builder.Writeln("Inserted paragraph while tracking.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Ensure a revision was created.
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
            Console.WriteLine("Revision accepted unexpectedly.");
        }
        catch (ObjectDisposedException ex)
        {
            Console.WriteLine("Caught expected exception: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Caught unexpected exception: " + ex.Message);
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
