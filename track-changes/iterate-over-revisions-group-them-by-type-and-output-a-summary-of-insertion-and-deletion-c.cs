using System;
using System.Collections.Generic;
using Aspose.Words;

namespace RevisionSummaryExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some initial text that will NOT be a revision.
            builder.Write("Hello world! ");

            // Start tracking revisions.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            // Insert new text – this will be recorded as an insertion revision.
            builder.Write("Inserted text. ");

            // Delete the original run – this will be recorded as a deletion revision.
            // The first run contains "Hello world! ".
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Count revisions by type.
            var revisionCounts = new Dictionary<RevisionType, int>();
            foreach (Revision rev in doc.Revisions)
            {
                if (revisionCounts.ContainsKey(rev.RevisionType))
                    revisionCounts[rev.RevisionType]++;
                else
                    revisionCounts[rev.RevisionType] = 1;
            }

            int insertionCount = revisionCounts.ContainsKey(RevisionType.Insertion) ? revisionCounts[RevisionType.Insertion] : 0;
            int deletionCount = revisionCounts.ContainsKey(RevisionType.Deletion) ? revisionCounts[RevisionType.Deletion] : 0;

            // Output the summary.
            Console.WriteLine($"Revision summary:");
            Console.WriteLine($"Insertions: {insertionCount}");
            Console.WriteLine($"Deletions: {deletionCount}");

            // Save the document (optional, demonstrates file output).
            doc.Save("RevisionsSummary.docx");
        }
    }
}
