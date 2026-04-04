using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document with a formatted comment.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");

        // Create a comment authored by Alice.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        // Ensure the comment contains at least one paragraph.
        comment.EnsureMinimum();

        // Add a run with bold formatting.
        Run commentRun = new Run(doc, "Original comment text.");
        commentRun.Font.Bold = true;
        comment.FirstParagraph?.AppendChild(commentRun);

        // Attach the comment to the first paragraph of the document body.
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);

        // Save the original document (optional, just for reference).
        doc.Save("Original.docx");

        // -----------------------------------------------------------------
        // 2. Update the text of the comment at a specific index while
        //    preserving its original formatting.
        // -----------------------------------------------------------------
        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        int targetIndex = 0; // zero‑based index of the comment to modify

        if (targetIndex >= 0 && targetIndex < comments.Count)
        {
            Comment targetComment = comments[targetIndex];

            // Ensure the comment has at least one paragraph.
            Paragraph? firstParagraph = targetComment.FirstParagraph;
            if (firstParagraph != null)
            {
                // Ensure there is at least one run to hold the text.
                if (firstParagraph.Runs.Count == 0)
                {
                    // Create an empty run with default formatting if none exist.
                    firstParagraph.AppendChild(new Run(doc, ""));
                }

                // Update the text of the first run. The run's Font (e.g., Bold) remains unchanged.
                Run firstRun = firstParagraph.Runs[0];
                firstRun.Text = "Updated comment text.";
                // Additional runs, if any, keep their original formatting untouched.
            }
        }

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save("Updated.docx");
    }
}
