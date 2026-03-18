using System;
using Aspose.Words;
using Aspose.Words.Tables;

class UpdateComment
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Index of the comment to update (zero‑based).
        int commentIndex = 0;

        // Retrieve the comment collection and the target comment.
        Comment comment = (Comment)doc.GetChildNodes(NodeType.Comment, true)[commentIndex];

        if (comment != null && comment.FirstParagraph != null && comment.FirstParagraph.Runs.Count > 0)
        {
            // Preserve the formatting of the first run.
            Run firstRun = comment.FirstParagraph.Runs[0];
            // Replace the text while keeping the existing Font, Style, etc.
            firstRun.Text = "Updated comment text preserving formatting.";

            // Remove any additional runs that may exist after the first one.
            for (int i = comment.FirstParagraph.Runs.Count - 1; i > 0; i--)
                comment.FirstParagraph.Runs[i].Remove();
        }
        else
        {
            // Fallback: use SetText if the comment has no runs (creates default formatting).
            comment.SetText("Updated comment text.");
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
