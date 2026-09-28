using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace TrackChangesDemo
{
    public class Program
    {
        // Returns true if any revision in the document was made by an author in the specified list.
        public static bool AnyRevisionAuthorMatches(Document doc, IEnumerable<string> authors)
        {
            var authorSet = new HashSet<string>(authors, StringComparer.OrdinalIgnoreCase);
            foreach (Revision rev in doc.Revisions)
            {
                if (authorSet.Contains(rev.Author))
                    return true;
            }
            return false;
        }

        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Original paragraph.");

            // First revision by Alice.
            doc.StartTrackRevisions("Alice", DateTime.Now);
            builder.Writeln("Alice adds this line.");
            doc.StopTrackRevisions();

            // Second revision by Bob.
            doc.StartTrackRevisions("Bob", DateTime.Now);
            builder.Writeln("Bob adds another line.");
            doc.StopTrackRevisions();

            // Save the document (optional, demonstrates file output).
            doc.Save("TrackedDocument.docx");

            // Define authors to check.
            var authorsToCheck = new List<string> { "Bob", "Charlie" };

            // Evaluate whether any revision matches the specified authors.
            bool hasMatch = AnyRevisionAuthorMatches(doc, authorsToCheck);

            // Output the result.
            Console.WriteLine($"Any revision authored by specified names? {hasMatch}");
        }
    }
}
