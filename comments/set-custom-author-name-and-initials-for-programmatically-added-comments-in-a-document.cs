using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will be associated with a comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment with custom author name and initials.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph and run with visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the paragraph we just added.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document to a file.
        string outputPath = "CommentWithCustomAuthor.docx";
        doc.Save(outputPath);

        // Enumerate all comments and output their metadata.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Initials: {c.Initial}, Text: {c.GetText().Trim()}");
        }
    }
}
