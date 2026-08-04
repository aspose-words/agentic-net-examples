using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This is a sample paragraph that will have a comment.");

        // Create a comment with custom author name and initials.
        Comment comment = new Comment(doc, "Jane Smith", "JS", DateTime.Now);
        // Set the visible text of the comment.
        comment.SetText("Please review this paragraph for accuracy.");

        // Attach the comment to the current paragraph.
        Paragraph? currentParagraph = builder.CurrentParagraph;
        if (currentParagraph != null)
        {
            currentParagraph.AppendChild(comment);
        }

        // Add a second comment to demonstrate multiple entries.
        Comment secondComment = new Comment(doc, "Bob Johnson", "BJ", DateTime.Now);
        secondComment.SetText("Consider rephrasing the first sentence.");
        if (currentParagraph != null)
        {
            currentParagraph.AppendChild(secondComment);
        }

        // Enumerate all comments and output their author, initials, and text.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author} ({c.Initial}): {c.GetText().Trim()}");
        }

        // Save the document to the working directory.
        doc.Save("CustomComment.docx");
    }
}
