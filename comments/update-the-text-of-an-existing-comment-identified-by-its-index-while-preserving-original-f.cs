using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");

        // Create a comment with some formatted text (bold).
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Original comment.");

        // Preserve the formatting of the first run (make it bold).
        Run? firstRun = comment.FirstParagraph?.Runs.OfType<Run>().FirstOrDefault();
        if (firstRun != null)
        {
            firstRun.Font.Bold = true;
        }

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // -----------------------------------------------------------------
        // Update the text of the comment at a specific index while preserving formatting.
        // -----------------------------------------------------------------
        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        int targetIndex = 0; // Index of the comment to update.

        if (targetIndex >= 0 && targetIndex < comments.Count)
        {
            Comment targetComment = comments[targetIndex];

            // Update the text of each run inside the comment's first paragraph.
            // This keeps the original formatting (e.g., bold, italic) intact.
            Paragraph? commentParagraph = targetComment.FirstParagraph;
            if (commentParagraph != null)
            {
                foreach (Run run in commentParagraph.Runs)
                {
                    run.Text = "Updated comment text.";
                }
            }
        }

        // Save the modified document.
        doc.Save("UpdatedComment.docx");
    }
}
