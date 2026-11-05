using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class PreserveCommentFormatting
{
    public static void Main()
    {
        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment with formatted text.");

        // Create a comment with bold and italic runs.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;
        Run italicRun = new Run(doc, " and italic text");
        italicRun.Font.Italic = true;

        commentParagraph.AppendChild(boldRun);
        commentParagraph.AppendChild(italicRun);
        comment.AppendChild(commentParagraph);

        // Attach the comment to the current paragraph.
        Paragraph? anchorParagraph = builder.CurrentParagraph;
        if (anchorParagraph != null)
        {
            anchorParagraph.AppendChild(comment);
        }

        // Make the comment content visible in the document by copying its formatted runs.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            // The comment is attached to a paragraph.
            Paragraph? parentPara = c.ParentNode as Paragraph;
            if (parentPara == null)
                continue;

            // The comment's first paragraph holds the formatted runs.
            Paragraph? sourcePara = c.FirstParagraph;
            if (sourcePara == null)
                continue;

            // Create a new paragraph that will hold the copied runs.
            Paragraph visiblePara = new Paragraph(doc);
            foreach (Node runNode in sourcePara.GetChildNodes(NodeType.Run, true))
            {
                visiblePara.AppendChild(runNode.Clone(true));
            }

            // Insert the new paragraph after the paragraph that contains the comment.
            Node? parentNode = parentPara.ParentNode;
            if (parentNode is CompositeNode compositeParent)
            {
                compositeParent.InsertAfter(visiblePara, parentPara);
            }
        }

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the document as HTML (comments are preserved with formatting).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        string htmlPath = Path.Combine(outputDir, "DocumentWithComments.html");
        doc.Save(htmlPath, htmlOptions);

        // Also save as DOCX for reference.
        string docxPath = Path.Combine(outputDir, "DocumentWithComments.docx");
        doc.Save(docxPath);
    }
}
