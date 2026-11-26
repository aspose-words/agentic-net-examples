using System;
using Aspose.Words;

public class AddCommentExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add two paragraphs to the document.
        builder.Writeln("First paragraph without comment.");
        builder.Writeln("Second paragraph that will receive a comment.");

        // Retrieve the second paragraph (index 1) from the document body.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.Paragraphs.Count > 1
            ? doc.FirstSection.Body.Paragraphs[1]
            : null;

        if (targetParagraph != null)
        {
            // Create a new comment.
            Comment comment = new Comment(doc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };

            // Add a paragraph and a run inside the comment so it has visible text.
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

            // Attach the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the document with the new comment.
        doc.Save("DocumentWithComment.docx");
    }
}
