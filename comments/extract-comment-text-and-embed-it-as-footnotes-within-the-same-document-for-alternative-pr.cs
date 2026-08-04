using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("First comment text.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment1);
        // Ensure the comment has a paragraph to hold its text.
        comment1.EnsureMinimum();
        // Move the builder into the comment and write its content.
        builder.MoveTo(comment1.FirstParagraph);
        builder.Write("Details of the first comment.");

        // Add second paragraph with a comment.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Second comment text.");
        builder.CurrentParagraph.AppendChild(comment2);
        comment2.EnsureMinimum();
        builder.MoveTo(comment2.FirstParagraph);
        builder.Write("Details of the second comment.");

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Environment.CurrentDirectory, "OriginalWithComments.docx");
        doc.Save(originalPath);

        // Extract all comments from the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Create a new document that will contain the footnotes.
        Document footnoteDoc = new Document();
        DocumentBuilder footnoteBuilder = new DocumentBuilder(footnoteDoc);

        footnoteBuilder.Writeln("Comments converted to footnotes:");
        footnoteBuilder.Writeln();

        // For each comment, write a line and insert a footnote with the comment text.
        foreach (Comment c in comments)
        {
            string author = c.Author ?? string.Empty;
            string commentText = c.GetText()?.Trim() ?? string.Empty;

            footnoteBuilder.Write($"Comment by {author}: ");
            footnoteBuilder.InsertFootnote(FootnoteType.Footnote, commentText);
            footnoteBuilder.Writeln();
        }

        // Save the document with footnotes.
        string footnotePath = Path.Combine(Environment.CurrentDirectory, "CommentsAsFootnotes.docx");
        footnoteDoc.Save(footnotePath);
    }
}
