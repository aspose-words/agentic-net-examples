using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace RevisionReportExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some initial content that will NOT be a revision.
            builder.Writeln("Original paragraph 1.");
            builder.Writeln("Original paragraph 2.");

            // Start tracking revisions with a specific author.
            doc.StartTrackRevisions("Alice", DateTime.Now);

            // Insert new paragraphs – these will be insertion revisions.
            builder.Writeln("Inserted paragraph A.");
            builder.Writeln("Inserted paragraph B.");

            // Modify an existing paragraph to create a deletion revision.
            // Remove the first run of the first original paragraph.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            if (firstParagraph.Runs.Count > 0)
                firstParagraph.Runs[0].Remove();

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Save the document (optional, just to have a file with revisions).
            doc.Save("RevisionsReport.docx");

            // Generate a report of each revision: type, author, and paragraph number.
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
                    : -1; // -1 indicates the revision is not attached to a paragraph.

                Console.WriteLine($"Revision {i + 1}:");
                Console.WriteLine($"  Type   : {rev.RevisionType}");
                Console.WriteLine($"  Author : {rev.Author}");
                Console.WriteLine($"  Paragraph #: {paragraphNumber}");
                Console.WriteLine();
            }
        }
    }
}
