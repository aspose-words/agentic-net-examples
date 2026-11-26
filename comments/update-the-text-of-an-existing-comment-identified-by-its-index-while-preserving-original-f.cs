using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class UpdateCommentExample
{
    public static void Main()
    {
        // Create a sample document with a formatted comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add a paragraph inside the comment.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Add a run with bold formatting inside the comment.
        Run commentRun = new Run(doc, "Original comment text.");
        commentRun.Font.Bold = true; // Apply bold formatting.
        commentParagraph.AppendChild(commentRun);

        // Attach the comment to the last paragraph of the main document.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the original document.
        string originalPath = "original.docx";
        doc.Save(originalPath);

        // Load the document (simulating a separate operation).
        Document loadedDoc = new Document(originalPath);

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Define the index of the comment to update.
        int commentIndex = 0; // Update the first comment.

        if (commentIndex >= 0 && commentIndex < comments.Count)
        {
            Comment commentToUpdate = comments[commentIndex];

            // Preserve original formatting by modifying the existing run's text.
            Paragraph? firstParagraph = commentToUpdate.FirstParagraph;
            if (firstParagraph != null)
            {
                Run? firstRun = firstParagraph.GetChildNodes(NodeType.Run, true)
                                             .OfType<Run>()
                                             .FirstOrDefault();
                if (firstRun != null)
                {
                    // Update the text while keeping the existing Font formatting.
                    firstRun.Text = "Updated comment text preserving formatting.";
                }
            }
        }

        // Save the updated document.
        string updatedPath = "updated.docx";
        loadedDoc.Save(updatedPath);
    }
}
