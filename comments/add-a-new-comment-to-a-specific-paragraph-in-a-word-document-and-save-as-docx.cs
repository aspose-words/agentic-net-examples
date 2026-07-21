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
        builder.Writeln("Second paragraph – this will receive a comment.");
        builder.Writeln("Third paragraph.");

        // Locate the second paragraph (index 1) in the body.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.Paragraphs.Count > 1
            ? doc.FirstSection.Body.Paragraphs[1]
            : null;

        if (targetParagraph != null)
        {
            // Create a new comment with author metadata.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("Review the content of this paragraph.");

            // Append the comment to the selected paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the document with the added comment.
        doc.Save("OutputWithComment.docx");
    }
}
