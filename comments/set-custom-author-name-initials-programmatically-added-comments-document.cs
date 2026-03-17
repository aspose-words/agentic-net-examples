using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This is a sample paragraph. ");

        // Create a comment with a custom author name and initials.
        // The constructor sets Author, Initial and DateTime in one step.
        Comment comment = new Comment(doc, "Alice Johnson", "AJ", DateTime.Now);
        comment.SetText("Review this sentence.");

        // Append the comment to the current paragraph.
        // The comment will appear in the right‑hand margin of the document.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to disk.
        doc.Save("CommentsWithCustomAuthor.docx");
    }
}
