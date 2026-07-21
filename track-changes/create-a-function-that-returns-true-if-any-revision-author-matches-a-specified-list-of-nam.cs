using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    // Returns true if any revision in the document was made by an author in the provided list.
    public static bool HasRevisionFromAuthors(Document doc, IEnumerable<string> authors)
    {
        // Use a HashSet for fast lookup.
        var authorSet = new HashSet<string>(authors);
        foreach (Revision rev in doc.Revisions)
        {
            if (authorSet.Contains(rev.Author))
                return true;
        }
        return false;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text (not a revision).
        builder.Writeln("Initial content. ");

        // First set of revisions by Alice.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Added by Alice. ");
        doc.StopTrackRevisions();

        // Second set of revisions by Charlie.
        doc.StartTrackRevisions("Charlie", DateTime.Now);
        builder.Writeln("Added by Charlie. ");
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates file output).
        doc.Save("Sample.docx");

        // Define authors to check.
        var authorsToCheck = new List<string> { "Bob", "Charlie" };

        // Evaluate whether any revision matches the specified authors.
        bool hasMatch = HasRevisionFromAuthors(doc, authorsToCheck);

        // Output the result.
        Console.WriteLine($"Document contains revision from specified authors: {hasMatch}");
    }
}
