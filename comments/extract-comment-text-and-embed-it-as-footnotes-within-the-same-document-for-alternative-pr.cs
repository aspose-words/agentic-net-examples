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
            // Create a new document and a builder for adding content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add first paragraph.
            builder.Writeln("First paragraph with a comment.");

            // Create a comment for the first paragraph.
            Comment comment1 = new Comment(doc)
            {
                Author = "Alice",
                Initial = "A",
                DateTime = DateTime.Now
            };
            // Add visible text to the comment.
            comment1.AppendChild(new Paragraph(doc));
            comment1.FirstParagraph?.AppendChild(new Run(doc, "Review the first paragraph."));

            // Attach the comment to the first paragraph node.
            Paragraph? firstPara = doc.FirstSection?.Body?.FirstParagraph;
            firstPara?.AppendChild(comment1);

            // Add second paragraph.
            builder.Writeln("Second paragraph with another comment.");

            // Create a comment for the second paragraph.
            Comment comment2 = new Comment(doc)
            {
                Author = "Bob",
                Initial = "B",
                DateTime = DateTime.Now
            };
            comment2.AppendChild(new Paragraph(doc));
            comment2.FirstParagraph?.AppendChild(new Run(doc, "Check the data in this paragraph."));

            // Attach the comment to the second paragraph node.
            Paragraph? secondPara = doc.FirstSection?.Body?.LastParagraph;
            secondPara?.AppendChild(comment2);

            // Save the original document (optional, for reference).
            doc.Save("original.docx");

            // Enumerate all comments in the document.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .ToList();

            // Move the builder to the end of the document to add a summary section.
            builder.MoveToDocumentEnd();
            builder.Writeln();
            builder.Writeln("Comments Summary:");

            // For each comment, write author info and insert a footnote with the comment text.
            foreach (Comment c in comments)
            {
                string commentText = c.GetText()?.Trim() ?? string.Empty;
                builder.Write($"Comment by {c.Author}: ");
                builder.InsertFootnote(FootnoteType.Footnote, commentText);
                builder.Writeln();
            }

            // Save the document with footnotes.
            doc.Save("comments-footnotes.docx");
        }
    }
}
