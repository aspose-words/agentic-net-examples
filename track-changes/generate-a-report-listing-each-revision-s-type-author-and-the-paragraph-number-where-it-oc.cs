using System;
using Aspose.Words;

namespace RevisionReportDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an initial paragraph – this will not be a revision.
            builder.Writeln("Original paragraph. This text will be partially deleted later.");

            // ---------- First set of revisions (author: Alice) ----------
            doc.StartTrackRevisions("Alice", DateTime.Now);

            // Insert a new paragraph – this will be recorded as an insertion revision.
            builder.Writeln("Inserted paragraph by Alice.");

            // Stop tracking for Alice.
            doc.StopTrackRevisions();

            // ---------- Second set of revisions (author: Bob) ----------
            doc.StartTrackRevisions("Bob", DateTime.Now);

            // Delete a run from the first paragraph – this will be recorded as a deletion revision.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            if (firstParagraph.Runs.Count > 0)
                firstParagraph.Runs[0].Remove();

            // Stop tracking for Bob.
            doc.StopTrackRevisions();

            // Save the document (optional, just to have an output file).
            doc.Save("RevisionsReport.docx");

            // Generate the report: list each revision's type, author, and paragraph number.
            Console.WriteLine("Revision Report:");
            Console.WriteLine("----------------");

            RevisionCollection revisions = doc.Revisions;
            for (int i = 0; i < revisions.Count; i++)
            {
                Revision rev = revisions[i];
                // Determine the paragraph that contains the revision.
                Paragraph revParagraph = rev.ParentNode?.GetAncestor(NodeType.Paragraph) as Paragraph;
                int paragraphNumber = revParagraph != null
                    ? doc.FirstSection.Body.Paragraphs.IndexOf(revParagraph) + 1
                    : -1; // Unknown

                Console.WriteLine($"Revision #{i + 1}");
                Console.WriteLine($"  Type   : {rev.RevisionType}");
                Console.WriteLine($"  Author : {rev.Author}");
                Console.WriteLine($"  Paragraph #: {(paragraphNumber > 0 ? paragraphNumber.ToString() : "N/A")}");
                Console.WriteLine();
            }
        }
    }
}
