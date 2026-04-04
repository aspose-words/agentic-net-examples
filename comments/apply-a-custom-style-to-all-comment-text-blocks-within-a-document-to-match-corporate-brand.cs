using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;

public class ApplyCommentStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample paragraphs.
        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph of the document.");

        // Insert the first comment.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Please review the first paragraph.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment1);

        // Insert the second comment.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Check the formatting of this line.");
        builder.Writeln("A line that needs attention.");
        // Append the comment to the newly added paragraph.
        builder.CurrentParagraph.AppendChild(comment2);

        // -----------------------------------------------------------------
        // Define a custom style that matches corporate branding.
        // -----------------------------------------------------------------
        const string styleName = "CorporateComment";
        // Add a new paragraph style to the document's style collection.
        Style corporateStyle = doc.Styles.Add(StyleType.Paragraph, styleName);
        // Configure the style's appearance (example: Arial, 12pt, dark blue, bold).
        corporateStyle.Font.Name = "Arial";
        corporateStyle.Font.Size = 12;
        corporateStyle.Font.Color = Color.DarkBlue;
        corporateStyle.Font.Bold = true;

        // -----------------------------------------------------------------
        // Apply the custom style to all comment text blocks.
        // -----------------------------------------------------------------
        var commentNodes = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

        foreach (Comment comment in commentNodes)
        {
            // Each comment contains its own collection of paragraphs.
            foreach (Paragraph paragraph in comment.Paragraphs)
            {
                // Apply the corporate style to the paragraph inside the comment.
                paragraph.ParagraphFormat.StyleName = styleName;
            }
        }

        // Save the modified document.
        const string outputPath = "CommentsStyled.docx";
        doc.Save(outputPath);
    }
}
