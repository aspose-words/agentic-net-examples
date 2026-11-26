using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Fonts;

public class HideCommentsExample
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment with visible text.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add a paragraph and run inside the comment so it has content.
        Paragraph commentParagraph = new Paragraph(doc);
        Run commentRun = new Run(doc, "Please review this paragraph.");
        commentParagraph.AppendChild(commentRun);
        comment.AppendChild(commentParagraph);

        // Attach the comment to the last paragraph of the main story.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the original document (comments are visible).
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);
        Console.WriteLine($"Original document saved to: {originalPath}");

        // Hide all comments by setting the hidden formatting on their runs.
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        foreach (Comment c in allComments)
        {
            var runs = c.GetChildNodes(NodeType.Run, true).OfType<Run>();
            foreach (Run r in runs)
            {
                // Set the hidden attribute; hidden text is not shown in the view.
                r.Font.Hidden = true;
            }
        }

        // Save the document with hidden comments.
        string hiddenPath = Path.Combine(outputDir, "hidden-comments.docx");
        doc.Save(hiddenPath);
        Console.WriteLine($"Document with hidden comments saved to: {hiddenPath}");
    }
}
