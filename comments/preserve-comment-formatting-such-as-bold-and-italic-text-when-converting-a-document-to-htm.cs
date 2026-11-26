using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // For CompositeNode

public class PreserveCommentFormatting
{
    public static void Main()
    {
        // Create a new document and a builder for adding content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will have a comment attached.
        builder.Writeln("This is a sample paragraph with a comment.");

        // Create a comment with bold and italic runs.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // The comment must contain at least one paragraph.
        comment.AppendChild(new Paragraph(doc));

        // Add a bold run.
        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;
        comment.FirstParagraph?.AppendChild(boldRun);

        // Add an italic run.
        Run italicRun = new Run(doc, " Italic text");
        italicRun.Font.Italic = true;
        comment.FirstParagraph?.AppendChild(italicRun);

        // Attach the comment to the last paragraph of the document.
        Paragraph targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // -----------------------------------------------------------------
        // Preserve comment formatting by copying the comment content into a
        // visible paragraph placed after the paragraph that contains the comment.
        // -----------------------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // The comment is anchored to a paragraph (its parent node).
            Paragraph anchorParagraph = c.ParentNode as Paragraph;
            if (anchorParagraph == null)
                continue;

            // Create a new paragraph that will hold the comment text.
            Paragraph visibleCommentParagraph = new Paragraph(doc);

            // Optional label for the comment.
            Run labelRun = new Run(doc, "Comment: ");
            labelRun.Font.Bold = true;
            visibleCommentParagraph.AppendChild(labelRun);

            // Copy each run from the comment's first paragraph, preserving formatting.
            Paragraph commentParagraph = c.FirstParagraph;
            if (commentParagraph != null)
            {
                foreach (Run run in commentParagraph.Runs.OfType<Run>())
                {
                    // Clone the run to keep its formatting.
                    Run clonedRun = (Run)run.Clone(true);
                    visibleCommentParagraph.AppendChild(clonedRun);
                }
            }

            // Insert the new paragraph after the anchor paragraph.
            CompositeNode parentNode = anchorParagraph.ParentNode as CompositeNode;
            if (parentNode != null)
            {
                parentNode.InsertAfter(visibleCommentParagraph, anchorParagraph);
            }
        }

        // Save the document as HTML. The visible comment paragraph retains bold/italic formatting.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.html");
        doc.Save(outputPath, saveOptions);
    }
}
