using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    // Returns true if any revision in the document was made by an author in the provided list.
    public static bool HasRevisionByAuthors(Document doc, List<string> authorNames)
    {
        foreach (Revision rev in doc.Revisions)
        {
            if (authorNames.Contains(rev.Author))
                return true;
        }
        return false;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First set of revisions by author "Alice".
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("First paragraph added by Alice.");
        doc.StopTrackRevisions();

        // Second set of revisions by author "Bob".
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Second paragraph added by Bob.");
        doc.StopTrackRevisions();

        // Save the document (optional for inspection).
        doc.Save("TrackedRevisions.docx");

        // Define authors to check.
        List<string> authorsToCheck = new List<string> { "Bob", "Charlie" };

        // Evaluate whether any revision matches the specified authors.
        bool hasMatchingRevision = HasRevisionByAuthors(doc, authorsToCheck);

        // Output the result.
        Console.WriteLine($"Document contains revision by specified author(s): {hasMatchingRevision}");
    }
}
