using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to simplify inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text that will be the anchor for the comment.
        builder.Write("Hello world!");

        // Create a top‑level comment with author, initials and date.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Today);

        // Append the comment to the current paragraph (the comment anchor).
        builder.CurrentParagraph.AppendChild(comment);

        // Move the builder into the comment's story to add the comment text.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
        builder.Write("This is a comment.");

        // Save the document to a DOCX file.
        doc.Save("CommentedDocument.docx");
    }
}
