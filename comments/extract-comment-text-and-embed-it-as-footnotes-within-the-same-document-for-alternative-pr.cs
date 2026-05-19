using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with some comments.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment.
        srcBuilder.Writeln("This is the first paragraph.");
        Paragraph firstPara = srcBuilder.CurrentParagraph;
        Comment comment1 = new Comment(sourceDoc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(sourceDoc));
        comment1.FirstParagraph?.AppendChild(new Run(sourceDoc, "Please review this sentence."));
        firstPara?.AppendChild(comment1);

        // Second paragraph with another comment.
        srcBuilder.Writeln("This is the second paragraph.");
        Paragraph secondPara = srcBuilder.CurrentParagraph;
        Comment comment2 = new Comment(sourceDoc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now.AddMinutes(-10)
        };
        comment2.AppendChild(new Paragraph(sourceDoc));
        comment2.FirstParagraph?.AppendChild(new Run(sourceDoc, "Consider rephrasing this part."));
        secondPara?.AppendChild(comment2);

        // Save the source document (optional, for inspection).
        sourceDoc.Save("source.docx");

        // Create a new document that will contain the comments as footnotes.
        Document footnoteDoc = new Document();
        DocumentBuilder footnoteBuilder = new DocumentBuilder(footnoteDoc);

        footnoteBuilder.Writeln("Comments converted to footnotes:");
        footnoteBuilder.Writeln();

        // Enumerate all comments in the source document.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            string author = c.Author ?? "Unknown";
            string text = c.GetText()?.Trim() ?? string.Empty;

            footnoteBuilder.Write($"Comment by {author}: ");
            footnoteBuilder.InsertFootnote(FootnoteType.Footnote, text);
            footnoteBuilder.Writeln();
        }

        // Save the document with footnotes.
        footnoteDoc.Save("comments-footnotes.docx");
    }
}
