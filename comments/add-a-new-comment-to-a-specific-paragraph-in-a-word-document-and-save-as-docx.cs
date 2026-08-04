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
        builder.Writeln("Second paragraph - target for comment.");
        builder.Writeln("Third paragraph.");

        // Retrieve the second paragraph (index 1) where the comment will be added.
        Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[1];

        // Create a comment with author, initials, and timestamp.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review this paragraph.");

        // Attach the comment to the selected paragraph.
        targetParagraph.AppendChild(comment);

        // Save the document as DOCX.
        doc.Save("CommentedDocument.docx");
    }
}
