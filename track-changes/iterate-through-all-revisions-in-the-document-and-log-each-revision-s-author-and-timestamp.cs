using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text (not tracked).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with the first author.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Added paragraph by Alice.");
        // Stop tracking so subsequent changes are not recorded as revisions for Alice.
        doc.StopTrackRevisions();

        // Start tracking revisions with a second author.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        // Delete the first paragraph to create a deletion revision.
        doc.FirstSection.Body.Paragraphs[0].Remove();
        // Add another paragraph.
        builder.Writeln("Added paragraph by Bob.");
        // Stop tracking.
        doc.StopTrackRevisions();

        // Save the document so the revisions are persisted.
        doc.Save("RevisionsDemo.docx");

        // Iterate through all revisions and log author and timestamp.
        foreach (Revision revision in doc.Revisions)
        {
            Console.WriteLine($"Author: {revision.Author}, Timestamp: {revision.DateTime}");
        }
    }
}
