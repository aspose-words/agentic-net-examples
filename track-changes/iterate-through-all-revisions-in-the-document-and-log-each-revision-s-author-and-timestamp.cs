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

        // Add some initial content that will not be a revision.
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions with a specific author and timestamp.
        string author = "Alice";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Make changes that will be recorded as revisions.
        builder.Writeln("First inserted line.");
        builder.Writeln("Second inserted line.");

        // Create a deletion revision by removing a run.
        // The first paragraph currently has one run ("This is the original paragraph.\r").
        // Remove that run to generate a deletion revision.
        if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
        {
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document so the revisions are persisted.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionsOutput.docx");
        doc.Save(outputPath);

        // Iterate through all revisions and log author and timestamp.
        foreach (Revision revision in doc.Revisions)
        {
            Console.WriteLine($"Revision Author: {revision.Author}, Timestamp: {revision.DateTime}");
        }
    }
}
