using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment with custom author name, initials, and the current date/time.
        Comment comment = new Comment(doc, "Alice Smith", "AS", DateTime.Now);
        comment.SetText("Please review this paragraph.");

        // Append the comment to the current paragraph if it exists.
        Paragraph? paragraph = builder.CurrentParagraph;
        if (paragraph != null)
        {
            paragraph.AppendChild(comment);
        }

        // Save the document to the working directory.
        const string outputPath = "CommentsSample.docx";
        doc.Save(outputPath);

        // Enumerate all comments in the document and output their metadata.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Initials: {c.Initial}, Text: {c.GetText().Trim()}");
        }
    }
}
