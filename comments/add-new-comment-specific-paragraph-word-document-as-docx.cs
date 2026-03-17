using System;
using Aspose.Words;

class AddCommentExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph – this will receive a comment.");
        builder.Writeln("Third paragraph.");

        // Retrieve the second paragraph (index 1) from the document body.
        Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[1];

        // Create a new comment anchored to the target paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Today);
        targetParagraph.AppendChild(comment);

        // Add a comment text paragraph inside the comment range.
        // Move the builder cursor to the newly created comment and insert a paragraph for the comment text.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
        builder.Write("This is the comment text.");

        // Save the document as DOCX.
        doc.Save("CommentedDocument.docx");
    }
}
