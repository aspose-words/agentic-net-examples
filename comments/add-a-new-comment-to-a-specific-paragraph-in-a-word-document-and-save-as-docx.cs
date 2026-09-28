using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph to comment.");
        builder.Writeln("Third paragraph.");

        // Locate the second paragraph (index 1) safely.
        Paragraph? targetParagraph = null;
        if (doc.FirstSection?.Body?.Paragraphs?.Count > 1)
        {
            targetParagraph = doc.FirstSection.Body.Paragraphs[1];
        }

        if (targetParagraph != null)
        {
            // Create a new comment and set its metadata.
            Comment comment = new Comment(doc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };

            // Add a paragraph and run inside the comment so it contains visible text.
            Paragraph commentParagraph = new Paragraph(doc);
            comment.AppendChild(commentParagraph);
            commentParagraph.AppendChild(new Run(doc, "This is a comment on the second paragraph."));

            // Attach the comment to the selected paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the modified document as a DOCX file.
        doc.Save("output.docx");
    }
}
