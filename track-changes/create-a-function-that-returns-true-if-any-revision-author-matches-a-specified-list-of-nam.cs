using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    // Checks if any revision in the document was made by an author in the provided list.
    public static bool HasRevisionFromAuthors(Document doc, List<string> authors)
    {
        foreach (Revision rev in doc.Revisions)
        {
            if (authors.Contains(rev.Author))
                return true;
        }
        return false;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content (this will not be a revision).
        builder.Writeln("Original text.");

        // Track changes made by Alice.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Alice's addition.");
        doc.StopTrackRevisions();

        // Track changes made by Bob.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Bob's addition.");

        // Create a deletion revision by removing the first paragraph.
        doc.FirstSection.Body.Paragraphs[0].Remove();
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates file output).
        doc.Save("RevisionsDemo.docx");

        // List of authors we want to check for.
        var authorsToCheck = new List<string> { "Charlie", "Bob" };

        // Use the helper function to determine if any matching revision exists.
        bool hasMatch = HasRevisionFromAuthors(doc, authorsToCheck);

        // Output the result.
        Console.WriteLine($"Has revision from specified authors: {hasMatch}");
    }
}
