using System;
using Aspose.Words;

class RevisionLogger
{
    // Records revision details to the console (could be extended to write to a file, database, etc.)
    public void Log(Revision revision)
    {
        Console.WriteLine($"Revision Type: {revision.RevisionType}, Author: {revision.Author}, Timestamp: {revision.DateTime}");
    }
}

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content that will not be tracked.
        builder.Writeln("Original paragraph.");

        // Begin tracking revisions with a specific author.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert new content – this will be recorded as an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Delete a run from the original paragraph – this will be recorded as a deletion revision.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Log each revision's type, author, and timestamp.
        RevisionLogger logger = new RevisionLogger();
        foreach (Revision rev in doc.Revisions)
        {
            logger.Log(rev);
        }

        // Save the document.
        doc.Save("RevisionsLogged.docx");
    }
}
