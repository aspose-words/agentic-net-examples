using System;
using System.Collections.Generic;
using Aspose.Words;

namespace RevisionAuthorCheck
{
    public class Program
    {
        // Returns true if any revision in the document has an author that exists in the provided list.
        public static bool HasRevisionByAuthors(Document doc, List<string> authors)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (authors == null) throw new ArgumentNullException(nameof(authors));

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

            // Write some initial text (not a revision).
            builder.Writeln("Initial content. ");

            // Start tracking revisions with author "John Doe".
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            builder.Writeln("Added by John. ");
            // Stop tracking for this author.
            doc.StopTrackRevisions();

            // Start tracking revisions with author "Jane Smith".
            doc.StartTrackRevisions("Jane Smith", DateTime.Now);
            builder.Writeln("Added by Jane. ");
            // Remove a run to create a deletion revision.
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
            doc.StopTrackRevisions();

            // Save the document so we can inspect it if needed.
            doc.Save("RevisionsDemo.docx");

            // Define a list of authors we are interested in.
            List<string> targetAuthors = new List<string> { "Jane Smith", "Alice Johnson" };

            // Check if any revision matches the target authors.
            bool hasMatchingRevision = HasRevisionByAuthors(doc, targetAuthors);

            // Output the result.
            Console.WriteLine($"Document contains revision by target author: {hasMatchingRevision}");
        }
    }
}
