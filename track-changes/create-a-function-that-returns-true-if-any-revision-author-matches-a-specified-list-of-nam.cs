using System;
using System.Collections.Generic;
using Aspose.Words;

namespace RevisionAuthorCheck
{
    public class Program
    {
        // Returns true if any revision in the document has an author that exists in the provided list.
        public static bool HasRevisionFromAuthors(Document doc, IList<string> authorNames)
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
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some initial text (not a revision).
            builder.Writeln("Initial content. ");

            // First set of revisions by author "Alice".
            doc.StartTrackRevisions("Alice", DateTime.Now);
            builder.Writeln("Alice's first revision. ");
            doc.StopTrackRevisions();

            // Second set of revisions by author "Bob".
            doc.StartTrackRevisions("Bob", DateTime.Now);
            builder.Writeln("Bob's revision. ");
            // Delete a run to create a deletion revision.
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
            doc.StopTrackRevisions();

            // Save the document (optional, demonstrates file output).
            doc.Save("Revisions.docx");

            // Define authors to check.
            List<string> authorsToCheck = new List<string> { "Bob", "Charlie" };

            // Perform the check.
            bool hasMatchingRevision = HasRevisionFromAuthors(doc, authorsToCheck);

            // Output the result.
            Console.WriteLine($"Document contains revision from specified authors: {hasMatchingRevision}");
        }
    }
}
