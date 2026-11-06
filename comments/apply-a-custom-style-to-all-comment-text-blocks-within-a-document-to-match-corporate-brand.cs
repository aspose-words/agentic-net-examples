using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

public class ApplyCommentStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph without a comment.
        builder.Writeln("First paragraph without comment.");

        // Add a paragraph that will have a comment.
        builder.Writeln("Second paragraph that will have a comment.");
        // Create the comment node.
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice Johnson",
            Initial = "AJ",
            DateTime = DateTime.Now
        };
        // Add a paragraph and run inside the comment to hold its visible text.
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review the data in this paragraph."));
        // Attach the comment to the paragraph that was just written.
        Paragraph? targetPara1 = builder.CurrentParagraph;
        targetPara1?.AppendChild(comment1);

        // Add another paragraph that will have a comment.
        builder.Writeln("Third paragraph with another comment.");
        // Create the second comment.
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob Smith",
            Initial = "BS",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        // Attach the second comment to its paragraph.
        Paragraph? targetPara2 = builder.CurrentParagraph;
        targetPara2?.AppendChild(comment2);

        // Save the original document for verification (optional).
        string originalPath = "OriginalDocument.docx";
        doc.Save(originalPath);

        // Define a custom style for comment text to match corporate branding.
        const string styleName = "CorporateCommentStyle";
        Style? existingStyle = doc.Styles[styleName];
        if (existingStyle == null)
        {
            Style newStyle = doc.Styles.Add(StyleType.Paragraph, styleName);
            newStyle.Font.Name = "Calibri";
            newStyle.Font.Size = 11;
            newStyle.Font.Color = System.Drawing.Color.DarkBlue;
            newStyle.Font.Bold = true;
        }

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Apply the custom style to each paragraph inside each comment.
        foreach (Comment comment in comments)
        {
            // Ensure the comment contains at least one paragraph.
            if (comment.FirstParagraph == null)
                continue;

            // Apply the style to every paragraph within the comment.
            foreach (Paragraph para in comment.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
            {
                para.ParagraphFormat.Style = doc.Styles[styleName];
            }
        }

        // Save the modified document.
        string outputPath = "DocumentWithStyledComments.docx";
        doc.Save(outputPath);
    }
}
