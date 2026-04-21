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

            // Add some initial paragraphs that are not tracked as revisions.
            builder.Writeln("Paragraph 1: Original text.");
            builder.Writeln("Paragraph 2: Original text.");
            builder.Writeln("Paragraph 3: Original text.");

            // Start tracking revisions with author "Alice".
            doc.StartTrackRevisions("Alice", DateTime.Now);

            // Insert a new paragraph (will be an insertion revision).
            builder.Writeln("Paragraph 4: Inserted while tracking.");

            // Modify text in the first paragraph (deletion of old run + insertion of new run).
            Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
            // Remove the existing run.
            firstPara.Runs[0].Remove();
            // Insert new run.
            builder.MoveTo(firstPara);
            builder.Write("Paragraph 1: Modified while tracking.");

            // Delete the second paragraph (will be a deletion revision).
            Paragraph secondPara = doc.FirstSection.Body.Paragraphs[1];
            secondPara.Remove();

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the document so that revisions are persisted.
            string docPath = "SampleRevisions.docx";
            doc.Save(docPath);

            // Generate a report of each revision: type, author, and paragraph number.
            Console.WriteLine("Revision Report:");
            Console.WriteLine("----------------");

            // Helper to find the paragraph number (1‑based) for a given node.
            int GetParagraphNumber(Node node)
            {
                // Ascend the node hierarchy until we find a Paragraph.
                while (node != null && !(node is Paragraph))
                {
                    node = node.ParentNode;
                }

                if (node is Paragraph para)
                {
                    // Find the index of the paragraph within the body.
                    ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
                    for (int i = 0; i < paragraphs.Count; i++)
                    {
                        if (paragraphs[i] == para)
                            return i + 1; // 1‑based numbering
                    }
                }

                // If we cannot determine a paragraph, return -1.
                return -1;
            }

            // Iterate through all revisions in the document.
            foreach (Revision rev in doc.Revisions)
            {
                // Determine the paragraph number where the revision occurs.
                int paraNumber = GetParagraphNumber(rev.ParentNode);

                // Prepare a readable revision type string.
                string revType = rev.RevisionType.ToString();

                // Output the revision details.
                Console.WriteLine($"Type: {revType}, Author: {rev.Author}, Paragraph #: {paraNumber}");
            }
        }
    }
}
