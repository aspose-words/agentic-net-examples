using System;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsLinqExample
{
    // Extension methods that enable LINQ queries directly on a Document.
    public static class DocumentExtensions
    {
        // Returns all Run nodes in the document.
        public static IEnumerable<Run> GetRuns(this Document doc)
        {
            // GetChildNodes returns a NodeList which implements IEnumerable<Node>.
            // Cast each Node to Run.
            return doc.GetChildNodes(NodeType.Run, true).Cast<Run>();
        }

        // Returns all Paragraph nodes in the document.
        public static IEnumerable<Paragraph> GetParagraphs(this Document doc)
        {
            return doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>();
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Populate the document with sample text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the first paragraph.");
            builder.Writeln("Replace this line with new text.");
            builder.Writeln("Another paragraph with the word old.");

            // -----------------------------------------------------------------
            // Example 1: Find all Run nodes that contain the word "old" and replace it.
            // -----------------------------------------------------------------
            var runsToReplace = doc.GetRuns()
                                   .Where(r => r.Text.IndexOf("old", StringComparison.OrdinalIgnoreCase) >= 0)
                                   .ToList();

            foreach (var run in runsToReplace)
            {
                // Simple case‑insensitive replace using ToLower for demonstration.
                run.Text = run.Text.Replace("old", "new", StringComparison.OrdinalIgnoreCase);
            }

            // -----------------------------------------------------------------
            // Example 2: Find all Paragraphs that contain the word "Replace"
            // and insert a new paragraph after each of them.
            // -----------------------------------------------------------------
            var matchingParagraphs = doc.GetParagraphs()
                                        .Where(p => p.GetText().Contains("Replace"))
                                        .ToList();

            foreach (var paragraph in matchingParagraphs)
            {
                // Create a new paragraph with custom text.
                Paragraph newParagraph = new Paragraph(doc);
                newParagraph.AppendChild(new Run(doc, "Inserted by LINQ query."));

                // Insert the new paragraph immediately after the matching one.
                paragraph.ParentNode.InsertAfter(newParagraph, paragraph);
            }

            // Save the modified document using the standard Save method.
            doc.Save("Result.docx");
        }
    }
}
