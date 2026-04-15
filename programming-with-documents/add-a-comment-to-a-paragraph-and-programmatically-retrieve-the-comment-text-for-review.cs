using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph that will have a comment.");

        // Create a comment with author information and set its text.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("This is a comment attached to the paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Retrieve and display all comment texts in the document.
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment c in comments)
        {
            // GetText returns the full comment content (including any replies).
            string commentText = c.GetText().Trim();
            Console.WriteLine($"Comment by {c.Author}: \"{commentText}\"");
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentExample.docx");
        doc.Save(outputPath);
    }
}
