using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    // Returns true if any revision in the document was made by an author whose name is in the provided list.
    public static bool AnyRevisionAuthorMatches(Document doc, IEnumerable<string> authorNames)
    {
        // Normalize author names for case‑insensitive comparison.
        var normalizedAuthors = new HashSet<string>(authorNames.Select(a => a?.Trim().ToLowerInvariant() ?? string.Empty));

        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Author != null && normalizedAuthors.Contains(rev.Author.Trim().ToLowerInvariant()))
                return true;
        }

        return false;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Enable tracking with author "Alice".
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Make some changes that will be recorded as revisions.
        builder.Writeln("Inserted paragraph while tracking.");
        builder.Writeln("Another inserted line.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Example usage of the function.
        var authorsToCheck = new List<string> { "Alice", "Bob" };
        bool hasMatchingAuthor = AnyRevisionAuthorMatches(doc, authorsToCheck);
        Console.WriteLine($"Any revision authored by Alice or Bob? {hasMatchingAuthor}");

        // Check with a list that does not contain the actual author.
        var otherAuthors = new List<string> { "Charlie", "Dave" };
        bool hasOtherAuthor = AnyRevisionAuthorMatches(doc, otherAuthors);
        Console.WriteLine($"Any revision authored by Charlie or Dave? {hasOtherAuthor}");

        // Save the document (optional, demonstrates file output).
        doc.Save("TrackedDocument.docx");
    }
}
