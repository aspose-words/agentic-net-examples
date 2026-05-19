using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Build a sample document with a formatted comment.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("This paragraph will have a comment attached.");

        // Create a comment and set its metadata.
        Comment comment = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };

        // Build the comment body: a paragraph with a bold run.
        Paragraph commentParagraph = new Paragraph(doc);
        Run commentRun = new Run(doc, "Original comment text.");
        commentRun.Font.Bold = true; // Preserve this formatting.
        commentParagraph.AppendChild(commentRun);
        comment.AppendChild(commentParagraph);

        // Attach the comment to the last paragraph of the main document body.
        doc.FirstSection.Body.LastParagraph.AppendChild(comment);

        // Save the initial document.
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Update the text of the first comment (index 0) while preserving formatting.
        // -------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        if (comments.Count > 0)
        {
            // Identify the comment by its index.
            Comment targetComment = comments[0];

            // Find the first run inside the comment (if any) and change its text.
            Run? firstRun = null;
            Paragraph? firstParagraph = targetComment.FirstParagraph;
            if (firstParagraph != null)
            {
                firstRun = firstParagraph.Runs.OfType<Run>().FirstOrDefault();
            }

            if (firstRun != null)
            {
                // Update only the textual content; formatting stays intact.
                firstRun.Text = "Updated comment text.";
            }
            else
            {
                // If the comment has no runs, replace its whole text.
                targetComment.SetText("Updated comment text.");
            }
        }

        // -------------------------------------------------
        // 3. Save the document after the update.
        // -------------------------------------------------
        string updatedPath = Path.Combine(outputDir, "updated.docx");
        doc.Save(updatedPath);
    }
}
