using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This is a sample paragraph.");

        // Create a comment anchored to the current paragraph.
        // Author = "Alice", Initials = "A", Date = now.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review this sentence.");

        // Append the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Retrieve the comment text programmatically.
        // GetText() returns the comment text plus a paragraph break, so trim it.
        string commentText = comment.GetText().Trim();
        Console.WriteLine("Comment text: " + commentText);

        // Example: iterate over all comments in the document.
        foreach (Comment c in doc.GetChildNodes(NodeType.Comment, true))
        {
            Console.WriteLine($"Author: {c.Author}, Text: {c.GetText().Trim()}");
        }

        // Save the document to disk.
        doc.Save("CommentExample.docx");
    }
}
