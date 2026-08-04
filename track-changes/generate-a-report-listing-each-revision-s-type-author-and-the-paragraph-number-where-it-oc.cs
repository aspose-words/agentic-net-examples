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

            // Add some initial content that will NOT be tracked as revisions.
            builder.Writeln("Paragraph 1: Original text.");
            builder.Writeln("Paragraph 2: Original text.");
            builder.Writeln("Paragraph 3: Original text.");

            // Start tracking revisions with a specific author.
            doc.StartTrackRevisions("Alice", DateTime.Now);

            // Insert a new paragraph – this will be an insertion revision.
            builder.Writeln("Paragraph 4: Inserted while tracking.");

            // Delete a run from the first paragraph – this will create a deletion revision.
            // Find the first paragraph and remove its first run.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            if (firstParagraph.Runs.Count > 0)
                firstParagraph.Runs[0].Remove();

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Save the document so you can inspect it manually if needed.
            const string outputPath = "RevisionReport.docx";
            doc.Save(outputPath);

            // Generate a report of each revision: type, author, and paragraph number.
            Console.WriteLine("Revision Report:");
            Console.WriteLine("----------------");

            // Iterate through all revisions in the document.
            foreach (Revision rev in doc.Revisions)
            {
                // Determine the paragraph that contains the revision's parent node.
                Node parent = rev.ParentNode;
                Paragraph revParagraph = (Paragraph)parent.GetAncestor(NodeType.Paragraph);

                // If the revision is not attached to a paragraph (unlikely for insert/delete), skip it.
                if (revParagraph == null)
                    continue;

                // Find the paragraph's index within the body (0‑based) and convert to 1‑based for reporting.
                int paragraphIndex = doc.FirstSection.Body.Paragraphs.IndexOf(revParagraph) + 1;

                // Output the revision details.
                Console.WriteLine($"Paragraph {paragraphIndex}: Type = {rev.RevisionType}, Author = {rev.Author}");
            }
        }
    }
}
