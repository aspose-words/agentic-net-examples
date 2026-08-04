using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will host the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a top‑level comment.
        Comment topComment = new Comment(doc, "Alice", "A", DateTime.Now);
        topComment.SetText("Please review this paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(topComment);

        // Add a reply to the top‑level comment.
        topComment.AddReply("Bob", "B", DateTime.Now, "I have reviewed it, looks good.");

        // Enumerate all comments to demonstrate the nesting (reply appears under its parent).
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // Replies have a non‑null Ancestor; indent them for clarity.
            string indent = c.Ancestor == null ? "" : "  ";
            Console.WriteLine($"{indent}Comment by {c.Author}: {c.GetText().Trim()}");
        }

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CommentWithReply.docx");
        doc.Save(outputPath);
    }
}
