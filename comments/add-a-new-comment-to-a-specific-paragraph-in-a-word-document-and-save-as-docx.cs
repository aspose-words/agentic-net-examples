using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph – target for the comment.");
        builder.Writeln("Third paragraph.");

        // Safely locate the second paragraph (index 1).
        Paragraph targetParagraph = null;
        if (doc.FirstSection?.Body?.Paragraphs?.Count > 1)
        {
            targetParagraph = doc.FirstSection.Body.Paragraphs[1];
        }

        if (targetParagraph == null)
        {
            // Paragraph not found – exit.
            return;
        }

        // Create a new comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a newly added comment.");

        // Append the comment to the target paragraph.
        targetParagraph.AppendChild(comment);

        // Save the document as DOCX in the working directory.
        doc.Save("CommentedDocument.docx");
    }
}
