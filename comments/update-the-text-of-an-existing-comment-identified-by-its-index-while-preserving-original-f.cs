using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment attached.");

        // Create a comment with some initial text.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Original comment text.");

        // Apply formatting (bold) to the comment text to demonstrate preservation.
        // FirstParagraph is guaranteed to exist after SetText, but we still check for null.
        Paragraph? commentParagraph = comment.FirstParagraph;
        if (commentParagraph != null)
        {
            // Runs collection may contain nodes of different types; filter to Run.
            Run? commentRun = commentParagraph.Runs.OfType<Run>().FirstOrDefault();
            if (commentRun != null)
            {
                commentRun.Font.Bold = true;
            }
        }

        // Attach the comment to the current paragraph.
        if (builder.CurrentParagraph != null)
        {
            builder.CurrentParagraph.AppendChild(comment);
        }

        // Save the document before the update (optional, just for demonstration).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Update the text of the existing comment while preserving formatting.
        // -----------------------------------------------------------------

        // Retrieve all comment nodes in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Ensure there is at least one comment.
        if (comments.Count == 0)
        {
            Console.WriteLine("No comments found to update.");
            return;
        }

        // Identify the comment by its index (e.g., the first comment, index 0).
        int targetIndex = 0;
        if (targetIndex < 0 || targetIndex >= comments.Count)
        {
            Console.WriteLine($"Comment index {targetIndex} is out of range.");
            return;
        }

        Comment targetComment = comments[targetIndex];

        // Preserve existing formatting by modifying the text of the existing run(s).
        Paragraph? targetParagraph = targetComment.FirstParagraph;
        if (targetParagraph != null)
        {
            Run? firstRun = targetParagraph.Runs.OfType<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                firstRun.Text = "Updated comment text while keeping formatting.";
            }
            else
            {
                // No runs found – replace the whole comment text.
                targetComment.SetText("Updated comment text.");
            }
        }
        else
        {
            // No paragraph inside the comment – replace the whole comment text.
            targetComment.SetText("Updated comment text.");
        }

        // Save the updated document.
        string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "updated.docx");
        doc.Save(updatedPath);
    }
}
