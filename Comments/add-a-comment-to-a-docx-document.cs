using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for convenient editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the document.
        builder.Write("Hello world!");

        // Create a top‑level comment.
        // Author = "John Doe", Initials = "JD", Date = today.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Today);

        // Set the comment's visible text.
        comment.SetText("This is a comment added via Aspose.Words.");

        // Append the comment to the current paragraph.
        // The comment will appear in the right‑hand margin of the document.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to a DOCX file.
        doc.Save("CommentAdded.docx");
    }
}
