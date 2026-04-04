using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add three paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph – this will receive a comment.");
        builder.Writeln("Third paragraph.");

        // Ensure the second paragraph (index 1) exists.
        if (doc.FirstSection?.Body?.Paragraphs?.Count > 1)
        {
            // Retrieve the second paragraph.
            Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[1];

            // Create a new comment with author metadata.
            Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
            comment.SetText("This is a comment attached to the second paragraph.");

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the modified document as DOCX in the working directory.
        doc.Save("CommentAdded.docx");
    }
}
