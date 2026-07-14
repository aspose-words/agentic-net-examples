using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a sample document with comments.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment.
        srcBuilder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(sourceDoc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(sourceDoc));
        comment1.FirstParagraph?.AppendChild(new Run(sourceDoc, "Comment on the first paragraph."));
        srcBuilder.CurrentParagraph?.AppendChild(comment1);

        // Second paragraph with a comment.
        srcBuilder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(sourceDoc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(sourceDoc));
        comment2.FirstParagraph?.AppendChild(new Run(sourceDoc, "Another comment, on the second paragraph."));
        srcBuilder.CurrentParagraph?.AppendChild(comment2);

        // Save the source document (optional, for verification).
        sourceDoc.Save("original.docx");

        // Create a new document that will contain footnotes derived from comments.
        Document footnoteDoc = new Document();
        DocumentBuilder footBuilder = new DocumentBuilder(footnoteDoc);

        footBuilder.Writeln("Comments converted to footnotes:");
        footBuilder.Writeln();

        // Enumerate all comments in the source document.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            string author = c.Author ?? "Unknown";
            string commentText = c.GetText()?.Trim() ?? string.Empty;

            footBuilder.Write($"Comment by {author}: ");
            footBuilder.InsertFootnote(FootnoteType.Footnote, commentText);
            footBuilder.Writeln();
        }

        // Save the document with footnotes.
        footnoteDoc.Save("comments-footnotes.docx");
    }
}
