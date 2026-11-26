using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs with comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice (allowed author).
        builder.Writeln("This is the first paragraph.");
        AddComment(doc, builder.CurrentParagraph, "Alice", "AL", "Please review this paragraph.");

        // Second paragraph with a comment from Charlie (to be rejected).
        builder.Writeln("This is the second paragraph.");
        AddComment(doc, builder.CurrentParagraph, "Charlie", "CH", "Consider rephrasing this sentence.");

        // Third paragraph with a comment from Bob (allowed author).
        builder.Writeln("This is the third paragraph.");
        AddComment(doc, builder.CurrentParagraph, "Bob", "BO", "Looks good to me.");

        // Save the original document.
        const string originalPath = "original.docx";
        doc.Save(originalPath);

        // Define which authors' comments should be kept.
        List<string> allowedAuthors = new List<string> { "Alice", "Bob" };

        // Retrieve all comments safely.
        List<Comment> allComments = doc.GetChildNodes(NodeType.Comment, true)
                                      .OfType<Comment>()
                                      .ToList();

        // Remove comments whose author is not in the allowed list.
        foreach (Comment comment in allComments)
        {
            if (!allowedAuthors.Contains(comment.Author, StringComparer.OrdinalIgnoreCase))
            {
                comment.Remove();
            }
        }

        // Save the revised document without the rejected comments.
        const string revisedPath = "revised.docx";
        doc.Save(revisedPath);
    }

    // Helper method to create and attach a comment to a given paragraph.
    private static void AddComment(Document doc, Paragraph? paragraph, string author, string initials, string commentText)
    {
        if (paragraph == null) return;

        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initials,
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph and run with visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, commentText));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the specified paragraph.
        paragraph.AppendChild(comment);
    }
}
