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
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Locate the second paragraph (index 1) safely.
        Paragraph? targetParagraph = null;
        ParagraphCollection? paragraphs = doc.FirstSection?.Body?.Paragraphs;

        if (paragraphs != null && paragraphs.Count > 1)
        {
            targetParagraph = paragraphs[1];
        }
        else if (paragraphs != null && paragraphs.Count > 0)
        {
            // Fallback to the first paragraph if the second does not exist.
            targetParagraph = paragraphs[0];
        }

        if (targetParagraph != null)
        {
            // Create a new comment with author metadata.
            Comment comment = new Comment(doc, "Alex", "AL", DateTime.Now);

            // Add visible text to the comment.
            comment.SetText("Review this paragraph.");

            // Attach the comment to the selected paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the document as DOCX in the working directory.
        doc.Save("output.docx");
    }
}
