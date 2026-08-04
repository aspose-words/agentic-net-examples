using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class UpdateCommentExample
{
    public static void Main()
    {
        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph that will contain a comment.");

        // Create a comment with some initial text and formatting.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Original comment text.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Define the index of the comment to update (0‑based).
        int commentIndex = 0;

        // Retrieve all comment nodes in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Ensure the requested index exists.
        if (commentIndex >= 0 && commentIndex < comments.Count)
        {
            Comment targetComment = comments[commentIndex];

            // Get the first paragraph inside the comment story.
            Paragraph commentParagraph = targetComment.FirstParagraph;
            if (commentParagraph != null)
            {
                // Preserve formatting by updating the text of the first run.
                if (commentParagraph.Runs.Count > 0)
                {
                    Run firstRun = commentParagraph.Runs[0];
                    firstRun.Text = "Updated comment text while preserving formatting.";

                    // Remove any additional runs that may exist.
                    for (int i = commentParagraph.Runs.Count - 1; i > 0; i--)
                        commentParagraph.Runs[i].Remove();
                }
                else
                {
                    // If there are no runs, simply add a new one.
                    commentParagraph.AppendChild(new Run(doc, "Updated comment text while preserving formatting."));
                }
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "UpdatedComment.docx");
        doc.Save(outputPath);
    }
}
