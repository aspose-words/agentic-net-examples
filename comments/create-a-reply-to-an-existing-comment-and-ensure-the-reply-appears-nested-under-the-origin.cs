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

        // Add a paragraph that will hold the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a top‑level comment.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Please review this paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph?.AppendChild(comment);

        // Add a reply to the comment. The reply will be nested under the original comment.
        comment.AddReply("Bob", "B", DateTime.Now, "I have reviewed it, looks good.");

        // Save the document so the comment and its reply are persisted.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentReplyExample.docx");
        doc.Save(outputPath);

        // Enumerate all comments to demonstrate the nesting.
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        foreach (Comment c in allComments)
        {
            // Top‑level comments have no ancestor.
            string level = c.Ancestor == null ? "Top‑level" : "Reply";
            Console.WriteLine($"{level} comment by {c.Author}: {c.GetText().Trim()}");
        }

        // Indicate where the file was saved.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
