using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph of text.
        builder.Writeln("This is a sample paragraph.");

        // Create a comment anchored to the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("This is a comment for review.");

        // Attach the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to the local file system.
        string filePath = "CommentExample.docx";
        doc.Save(filePath);

        // Retrieve and display the comment text programmatically.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment c in commentNodes)
        {
            // Get the full text of the comment (including any child paragraphs).
            string commentText = c.GetText().Trim();
            Console.WriteLine($"Comment by {c.Author}: {commentText}");
        }
    }
}
