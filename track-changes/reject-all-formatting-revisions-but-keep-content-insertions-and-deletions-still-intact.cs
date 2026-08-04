using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will later be modified.
        builder.Write("Hello ");
        builder.Write("World");

        // Start tracking revisions.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Write("Inserted ");

        // Apply a formatting change to the first run.
        // (Aspose.Words currently does not record formatting changes as revisions,
        // but we include this step to illustrate the intended workflow.)
        Run firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Font.Bold = true;

        // Delete the original "World" text – this will be recorded as a deletion revision.
        Run worldRun = doc.FirstSection.Body.FirstParagraph.Runs[2];
        worldRun.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only formatting revisions, leaving insertions and deletions untouched.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
            {
                rev.Reject();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionsResult.docx");
        doc.Save(outputPath);

        // Output the final document text to the console for verification.
        Console.WriteLine("Final document text:");
        Console.WriteLine(doc.GetText());
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
