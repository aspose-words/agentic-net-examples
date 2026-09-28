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

        // Start tracking revisions with a specific author and timestamp.
        string author = "John Doe";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Insert a new paragraph – this will be recorded as an insertion revision.
        builder.Writeln("Inserted paragraph while tracking.");

        // Delete the original paragraph – this will be recorded as a deletion revision.
        Paragraph originalParagraph = doc.FirstSection.Body.Paragraphs[0];
        originalParagraph.Remove();

        // Change formatting of the inserted paragraph – this will be recorded as a format change revision.
        Paragraph insertedParagraph = doc.FirstSection.Body.Paragraphs[0];
        if (insertedParagraph.Runs.Count > 0)
        {
            insertedParagraph.Runs[0].Font.Size = 16;
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional for this example).
        doc.Save("Revisions.docx");

        // Iterate through all revisions and log author and timestamp.
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
