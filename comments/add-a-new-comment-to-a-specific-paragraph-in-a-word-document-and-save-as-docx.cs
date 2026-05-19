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

        // Retrieve the second paragraph (index 1) from the document body.
        Paragraph? secondParagraph = doc.FirstSection?.Body?.Paragraphs.Count > 1
            ? doc.FirstSection.Body.Paragraphs[1]
            : null;

        if (secondParagraph != null)
        {
            // Create a new comment with author metadata.
            Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
            comment.SetText("Review the wording of this paragraph.");

            // Append the comment to the selected paragraph.
            secondParagraph.AppendChild(comment);
        }

        // Save the modified document as DOCX.
        const string outputFile = "CommentAdded.docx";
        doc.Save(outputFile);
    }
}
