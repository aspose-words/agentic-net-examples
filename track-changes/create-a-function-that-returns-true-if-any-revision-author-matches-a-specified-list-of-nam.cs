using System;
using System.Collections.Generic;
using Aspose.Words;

public class RevisionAuthorChecker
{
    // Checks whether any revision in the document was made by one of the specified authors.
    public static bool HasRevisionByAuthors(Document doc, IList<string> authorNames)
    {
        // Iterate through all revisions in the document.
        foreach (Revision rev in doc.Revisions)
        {
            // If the revision's author matches any name in the list, return true.
            if (authorNames.Contains(rev.Author))
                return true;
        }
        // No matching author found.
        return false;
    }

    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text (not a revision).
        builder.Writeln("Initial content. ");

        // Start tracking revisions with author "Alice".
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Added by Alice. ");
        // Stop tracking for Alice.
        doc.StopTrackRevisions();

        // Start tracking revisions with author "Bob".
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Added by Bob. ");
        // Delete a run to create a deletion revision (still by Bob).
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        // Stop tracking for Bob.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates persistence).
        doc.Save("RevisionsExample.docx");

        // Define a list of authors to check against.
        List<string> authorsToFind = new List<string> { "Bob", "Charlie" };

        // Use the helper function to determine if any revision matches the list.
        bool hasMatchingRevision = HasRevisionByAuthors(doc, authorsToFind);

        // Output the result.
        Console.WriteLine($"Document contains revision by specified author(s): {hasMatchingRevision}");
    }
}
