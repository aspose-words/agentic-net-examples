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

            // Add sample paragraphs and attach comments to them.
            builder.Writeln("First paragraph with a comment.");
            Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
            comment1.SetText("This is Alice's comment on the first paragraph.");
            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment1);

            builder.Writeln("Second paragraph with another comment.");
            Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
            comment2.SetText("Bob's comment provides additional insight.");
            builder.CurrentParagraph.AppendChild(comment2);

            builder.Writeln("Third paragraph without a comment.");

            // Enumerate all comment nodes in the document.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            // For each comment, insert a footnote containing the comment's text.
            foreach (Comment c in comments)
            {
                // The comment should be anchored to a paragraph.
                Paragraph? parentParagraph = c.ParentParagraph;
                if (parentParagraph == null)
                    continue; // Skip if the comment is not attached to a paragraph.

                // Move the builder to the paragraph that contains the comment.
                builder.MoveTo(parentParagraph);
                // Insert a footnote with the plain text of the comment.
                builder.InsertFootnote(FootnoteType.Footnote, c.GetText().Trim());
            }

            // Save the resulting document.
            const string outputPath = "CommentsToFootnotes.docx";
            doc.Save(outputPath);
        }
    }
}
