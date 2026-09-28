using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph with a comment.");

        // Create a comment containing formatted (bold) text.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        Paragraph commentParagraph = new Paragraph(doc);
        Run commentRun = new Run(doc, "Original comment text.");
        commentRun.Font.Bold = true; // formatting to be preserved.
        commentParagraph.AppendChild(commentRun);
        comment.AppendChild(commentParagraph);

        // Attach the comment to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);

        // Save the original document.
        doc.Save("original.docx");

        // Load the document back (simulating a separate operation).
        Document loadedDoc = new Document("original.docx");

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Update the text of the comment at index 0 while preserving its formatting.
        if (comments.Count > 0)
        {
            Comment targetComment = comments[0];

            // Get the first paragraph inside the comment.
            Paragraph? firstParagraph = targetComment.FirstParagraph;
            if (firstParagraph != null && firstParagraph.HasChildNodes)
            {
                // Find the first Run node inside that paragraph.
                Run? firstRun = firstParagraph.GetChildNodes(NodeType.Run, true)
                    .OfType<Run>()
                    .FirstOrDefault();

                if (firstRun != null)
                {
                    // Replace the text; formatting (e.g., Bold) remains unchanged.
                    firstRun.Text = "Updated comment text.";
                }
            }
        }

        // Save the document with the updated comment.
        loadedDoc.Save("updated.docx");
    }
}
