using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

namespace CommentToFootnoteExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add first paragraph and attach a comment.
            builder.Writeln("First paragraph of the document.");
            Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
            comment1.SetText("Review the wording of this paragraph.");
            builder.CurrentParagraph?.AppendChild(comment1);

            // Add second paragraph and attach another comment.
            builder.Writeln("Second paragraph with additional content.");
            Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
            comment2.SetText("Consider adding a reference here.");
            builder.CurrentParagraph?.AppendChild(comment2);

            // Save the original document (optional, for inspection).
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string originalPath = Path.Combine(outputDir, "original.docx");
            doc.Save(originalPath);

            // Insert a summary section that converts each comment into a footnote.
            builder.Writeln(); // Add a blank line before the summary.
            builder.Writeln("Comments Summary (converted to footnotes):");

            // Enumerate all comments safely.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                // Extract plain comment text.
                string commentText = c.GetText().Trim();

                // Write a label and insert the footnote containing the comment text.
                builder.Write($"Comment by {c.Author}: ");
                builder.InsertFootnote(FootnoteType.Footnote, commentText);
                builder.Writeln(); // Move to the next line.
            }

            // Save the document with footnotes.
            string resultPath = Path.Combine(outputDir, "comments-footnotes.docx");
            doc.Save(resultPath);
        }
    }
}
