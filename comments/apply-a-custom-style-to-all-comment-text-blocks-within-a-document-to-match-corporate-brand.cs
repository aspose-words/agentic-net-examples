using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph where a comment will be added.");

        // Create a comment and attach it to the second paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "Jane Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add a paragraph and run inside the comment to hold visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph for accuracy."));
        comment.AppendChild(commentParagraph);

        // Append the comment to the paragraph.
        Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[1];
        targetParagraph.AppendChild(comment);

        // Create another comment for demonstration.
        Comment secondComment = new Comment(doc)
        {
            Author = "John Smith",
            Initial = "JS",
            DateTime = DateTime.Now
        };
        Paragraph secondCommentParagraph = new Paragraph(doc);
        secondCommentParagraph.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        secondComment.AppendChild(secondCommentParagraph);
        doc.FirstSection.Body.Paragraphs[0].AppendChild(secondComment);

        // Save the original document (optional, for reference).
        string originalPath = "OriginalDocument.docx";
        doc.Save(originalPath);

        // Define a custom style that matches corporate branding.
        Style corporateStyle = doc.Styles.Add(StyleType.Paragraph, "CorporateComment");
        corporateStyle.Font.Name = "Arial";
        corporateStyle.Font.Size = 10;
        corporateStyle.Font.Color = System.Drawing.Color.DarkBlue;
        corporateStyle.Font.Italic = true;

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Apply the custom style to each paragraph inside each comment.
        foreach (Comment c in comments)
        {
            // Ensure the comment has at least one paragraph.
            foreach (Paragraph p in c.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
            {
                // Apply the custom style by name.
                p.ParagraphFormat.StyleName = corporateStyle.Name;
            }
        }

        // Save the modified document.
        string outputPath = "DocumentWithStyledComments.docx";
        doc.Save(outputPath);
    }
}
