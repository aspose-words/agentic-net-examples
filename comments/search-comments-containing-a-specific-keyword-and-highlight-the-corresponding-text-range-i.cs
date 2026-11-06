using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for constructing its content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph without any comment.
        builder.Writeln("This paragraph has no comments.");

        // Paragraph that will have a comment containing the keyword "important".
        builder.Writeln("This paragraph will be annotated.");
        // The paragraph we just wrote is the current one.
        Paragraph? paraWithKeyword = builder.CurrentParagraph;
        if (paraWithKeyword != null)
        {
            // Create the comment.
            Comment comment = new Comment(doc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };
            // Add visible text to the comment body.
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph?.AppendChild(new Run(doc, "This is an important comment."));
            // Attach the comment to the paragraph.
            paraWithKeyword.AppendChild(comment);
        }

        // Paragraph that will have a comment without the keyword.
        builder.Writeln("Another paragraph with a regular comment.");
        Paragraph? paraWithoutKeyword = builder.CurrentParagraph;
        if (paraWithoutKeyword != null)
        {
            Comment otherComment = new Comment(doc)
            {
                Author = "Jane Smith",
                Initial = "JS",
                DateTime = DateTime.Now
            };
            otherComment.AppendChild(new Paragraph(doc));
            otherComment.FirstParagraph?.AppendChild(new Run(doc, "Just a regular note."));
            paraWithoutKeyword.AppendChild(otherComment);
        }

        // Define the keyword to search for inside comment texts.
        const string keyword = "important";

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // Check if the comment text contains the keyword (case‑insensitive).
            string commentText = c.GetText();
            if (!string.IsNullOrEmpty(commentText) &&
                commentText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // The comment is attached to a paragraph; highlight that paragraph.
                if (c.ParentNode is Paragraph para)
                {
                    foreach (Run run in para.Runs)
                    {
                        run.Font.HighlightColor = Color.Yellow;
                    }
                }
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "HighlightedComments.docx");
        doc.Save(outputPath);
    }
}
