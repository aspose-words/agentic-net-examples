using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add some paragraphs to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // -------------------------------------------------
        // Add a comment anchored to the first paragraph.
        // -------------------------------------------------
        // Create the comment node with author information.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Today);

        // The comment must be a child of a Paragraph node.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.AppendChild(comment);

        // Add the comment's visible text inside a new paragraph.
        Paragraph commentParagraph = (Paragraph)comment.AppendChild(new Paragraph(doc));
        commentParagraph.AppendChild(new Run(doc, "This is a comment."));

        // -------------------------------------------------
        // Add a reply to the previously created comment.
        // -------------------------------------------------
        comment.AddReply("Bob", "B", DateTime.Today, "Reply to comment.");

        // -------------------------------------------------
        // Save the document using the provided Save method.
        // -------------------------------------------------
        doc.Save("CommentsDemo.docx");
    }
}
