using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment with custom author name and initials.
        Comment comment = new Comment(doc)
        {
            Author = "Jane Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph and run to be visible.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Append the comment to the first paragraph of the document body.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the document to the working directory.
        string outputPath = "CommentWithCustomAuthor.docx";
        doc.Save(outputPath);

        // Enumerate all comments and write their metadata to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Initials: {c.Initial}, Text: {c.GetText().Trim()}");
        }
    }
}
