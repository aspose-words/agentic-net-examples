using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("Paragraph that will have a comment.");

        // Create a top‑level comment.
        Comment topComment = new Comment(doc, "Alice", "A", DateTime.Now);
        topComment.SetText("This is the original comment.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(topComment);

        // Add a reply to the top‑level comment.
        topComment.AddReply("Bob", "B", DateTime.Now, "This is a reply nested under the original comment.");

        // Save the document so the comment and its reply are persisted.
        doc.Save("CommentWithReply.docx");

        // Optional: enumerate comments to verify the hierarchy.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>())
        {
            string indent = comment.Ancestor == null ? "" : "    ";
            Console.WriteLine($"{indent}Author: {comment.Author}");
            Console.WriteLine($"{indent}Text: {comment.GetText().Trim()}");
        }
    }
}
