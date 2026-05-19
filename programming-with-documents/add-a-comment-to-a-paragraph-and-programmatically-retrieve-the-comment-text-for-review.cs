using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph to the document.
        builder.Writeln("This is a sample paragraph.");

        // Create a comment attached to the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Please review this paragraph.");

        // Append the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to a file.
        const string fileName = "CommentExample.docx";
        doc.Save(fileName);

        // Retrieve all comments from the document.
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment c in comments)
        {
            // Output the author and the comment text.
            Console.WriteLine($"Author: {c.Author}");
            Console.WriteLine($"Comment: {c.GetText().Trim()}");
        }
    }
}
