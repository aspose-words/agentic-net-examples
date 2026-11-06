using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class UpdateCommentExample
{
    public static void Main()
    {
        // Create a temporary folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Build a sample document that contains a comment.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment with some formatted text (bold).
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Add a run with bold formatting inside the comment.
        Run commentRun = new Run(doc, "Original comment text.");
        commentRun.Font.Bold = true;
        commentParagraph.AppendChild(commentRun);

        // Attach the comment to the first paragraph of the document body.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document and update the comment text at index 0.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Retrieve all comment nodes.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Ensure the requested index exists.
        int targetIndex = 0;
        if (targetIndex >= 0 && targetIndex < comments.Count)
        {
            Comment targetComment = comments[targetIndex];

            // New text that will replace the existing comment content.
            string newCommentText = "Updated comment text preserving formatting.";

            // Update each Run inside the comment's first paragraph.
            // This keeps the original formatting (e.g., bold, italic) intact.
            Paragraph? commentFirstParagraph = targetComment.FirstParagraph;
            if (commentFirstParagraph != null)
            {
                var runs = commentFirstParagraph.GetChildNodes(NodeType.Run, true)
                                                .OfType<Run>()
                                                .ToList();

                foreach (Run run in runs)
                {
                    run.Text = newCommentText;
                }
            }
        }

        // Save the modified document.
        string updatedPath = Path.Combine(outputDir, "updated.docx");
        loadedDoc.Save(updatedPath);
    }
}
