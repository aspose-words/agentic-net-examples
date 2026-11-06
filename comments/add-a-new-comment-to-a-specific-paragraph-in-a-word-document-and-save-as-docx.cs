using System;
using Aspose.Words;

public class AddCommentExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph that will receive a comment.");
        builder.Writeln("Third paragraph.");

        // Retrieve the second paragraph (index 1) where the comment will be added.
        Paragraph targetParagraph = doc.FirstSection?.Body?.Paragraphs[1];
        if (targetParagraph == null)
        {
            throw new InvalidOperationException("Target paragraph not found.");
        }

        // Create a new comment with author metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add visible text to the comment body.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the target paragraph.
        targetParagraph.AppendChild(comment);

        // Save the modified document as DOCX.
        doc.Save("DocumentWithComment.docx");
    }
}
