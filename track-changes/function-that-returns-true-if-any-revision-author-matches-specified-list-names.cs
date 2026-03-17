using System;
using System.Collections.Generic;
using Aspose.Words;

public static class RevisionHelper
{
    // Returns true if any revision in the document has an author that exists in the supplied list.
    public static bool AnyRevisionAuthorMatches(Document doc, IEnumerable<string> authorNames)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (authorNames == null) throw new ArgumentNullException(nameof(authorNames));

        // Convert the list of names to a HashSet for fast lookup.
        var authorSet = new HashSet<string>(authorNames);

        // Iterate through all revisions in the document.
        foreach (Revision revision in doc.Revisions)
        {
            // If the revision's author is in the set, we have a match.
            if (authorSet.Contains(revision.Author))
                return true;
        }

        // No matching author found.
        return false;
    }
}

// Minimal console entry point required for a .NET console project.
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage – replace the path with a real .docx file that contains revisions.
        // The example is kept simple; it only demonstrates that the project now builds.
        // In real scenarios you would load a document and pass a list of author names.
        try
        {
            var doc = new Document("SampleWithRevisions.docx");
            var authorsToCheck = new List<string> { "Alice", "Bob" };
            bool hasMatch = RevisionHelper.AnyRevisionAuthorMatches(doc, authorsToCheck);
            Console.WriteLine($"Any matching revision author: {hasMatch}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
