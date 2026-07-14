using System;
using Aspose.Words;

namespace RevisionReportExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an initial paragraph – this will not be a revision.
            builder.Writeln("Original paragraph.");

            // Start tracking revisions with a specific author.
            doc.StartTrackRevisions("Alice", DateTime.Now);

            // Insert a new paragraph – this will be recorded as an insertion revision.
            builder.Writeln("Inserted paragraph.");

            // Append some text to the first paragraph – also an insertion revision.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            Run run = new Run(doc, " Added text.");
            firstParagraph.Runs.Add(run);

            // Delete a run from the first paragraph – creates a deletion revision.
            // Remove the word "Original" (first run) if it exists.
            if (firstParagraph.Runs.Count > 0)
                firstParagraph.Runs[0].Remove();

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Save the document (optional, just to demonstrate saving).
            string docPath = "RevisionsReport.docx";
            doc.Save(docPath);

            // Generate a report of revisions: type, author, and paragraph number.
            Console.WriteLine("Revision Type\tAuthor\tParagraph Number");
            foreach (Revision rev in doc.Revisions)
            {
                // Determine the paragraph that contains the revision.
                Paragraph paragraph = null;
                Node parentNode = rev.ParentNode;

                // Revision may be attached directly to a paragraph.
                if (parentNode is Paragraph para)
                {
                    paragraph = para;
                }
                // Revision may be attached to a run; use its ParentParagraph.
                else if (parentNode is Run runNode)
                {
                    paragraph = runNode.ParentParagraph;
                }
                // Fallback: try to climb one level up to find a paragraph.
                else
                {
                    paragraph = parentNode?.ParentNode as Paragraph;
                }

                // If we could not locate a paragraph, skip this revision.
                if (paragraph == null)
                    continue;

                // Find the paragraph's index (1‑based).
                int paragraphNumber = doc.FirstSection.Body.Paragraphs.IndexOf(paragraph) + 1;

                // Output the revision details.
                Console.WriteLine($"{rev.RevisionType}\t{rev.Author}\t{paragraphNumber}");
            }
        }
    }
}
