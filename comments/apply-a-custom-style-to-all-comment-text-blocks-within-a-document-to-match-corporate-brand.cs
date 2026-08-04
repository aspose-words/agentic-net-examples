using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample paragraphs.
        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph of the document.");

        // Insert a comment on the first paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Review the first paragraph for clarity.");
        // Append the comment to the paragraph.
        builder.MoveToDocumentStart();
        builder.CurrentParagraph.AppendChild(comment1);

        // Insert a second comment on the second paragraph.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing this sentence.");
        builder.MoveToDocumentEnd();
        builder.CurrentParagraph.AppendChild(comment2);

        // Define a custom style that matches corporate branding.
        const string styleName = "CorporateComment";
        Style corporateStyle = doc.Styles.Add(StyleType.Paragraph, styleName);
        corporateStyle.Font.Name = "Arial";
        corporateStyle.Font.Size = 10;
        corporateStyle.Font.Color = System.Drawing.Color.DarkBlue;
        corporateStyle.Font.Bold = true;
        corporateStyle.Font.Italic = false;

        // Apply the custom style to every comment in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            // Ensure the comment has at least one paragraph.
            comment.EnsureMinimum();

            // Apply the style to each paragraph inside the comment.
            foreach (Paragraph para in comment.Paragraphs)
            {
                para.ParagraphFormat.Style = corporateStyle;
            }
        }

        // Save the resulting document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CommentStyled.docx");
        doc.Save(outputPath);
    }
}
