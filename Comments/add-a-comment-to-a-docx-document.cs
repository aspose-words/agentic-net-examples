using System;
using Aspose.Words;

class AddCommentExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text that will host the comment.
        builder.Write("Hello world!");

        // Create a top‑level comment with author information.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Today);

        // Set the comment's visible text.
        comment.SetText("This is a comment attached to the paragraph.");

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Move the builder inside the comment to add additional paragraphs if needed.
        // Here we add a simple paragraph with the same text as the comment.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
        builder.Write("This is a comment attached to the paragraph.");

        // Save the document to a DOCX file.
        doc.Save("CommentAdded.docx");
    }
}
