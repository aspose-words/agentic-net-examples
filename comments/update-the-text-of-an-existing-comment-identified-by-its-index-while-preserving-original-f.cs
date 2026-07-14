using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment.");

        // Create a comment with some initial text and formatting.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Original comment text.");

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // -----------------------------------------------------------------
        // Update the text of the first comment (index 0) while preserving its formatting.
        // -----------------------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        if (comments.Count > 0)
        {
            Comment targetComment = comments[0];

            // The comment body consists of paragraphs and runs.
            // Preserve existing formatting by modifying the text of the first run.
            Paragraph? firstParagraph = targetComment.FirstParagraph;
            if (firstParagraph != null && firstParagraph.Runs.Count > 0)
            {
                // Replace the text of the first run; its formatting remains unchanged.
                firstParagraph.Runs[0].Text = "Updated comment text.";
            }
            else
            {
                // Fallback: if the comment has no runs, replace the whole text.
                targetComment.SetText("Updated comment text.");
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedComment.docx");
        doc.Save(outputPath);
    }
}
