using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;
using System.Drawing;

public class ApplyStyleToComments
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This is the first paragraph with a comment.");

        // Create the first comment and attach it to the current paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Please review this paragraph for accuracy.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Add another paragraph with a second comment.
        builder.Writeln("Second paragraph, also needs attention.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing this sentence.");
        builder.CurrentParagraph.AppendChild(comment2);

        // -----------------------------------------------------------------
        // Define a custom style that matches corporate branding for comments.
        // -----------------------------------------------------------------
        const string corporateStyleName = "CorporateComment";
        Style corporateStyle = doc.Styles.Add(StyleType.Paragraph, corporateStyleName);
        corporateStyle.Font.Name = "Arial";
        corporateStyle.Font.Size = 11;
        corporateStyle.Font.Color = Color.DarkBlue;
        corporateStyle.Font.Bold = true;
        corporateStyle.Font.Italic = false;

        // ---------------------------------------------------------------
        // Apply the custom style to every paragraph inside each comment.
        // ---------------------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList(); // Safe copy for iteration.

        foreach (Comment comment in comments)
        {
            // Ensure the comment contains at least one paragraph.
            if (comment.FirstParagraph == null)
                continue;

            // Apply the style to each paragraph within the comment.
            foreach (Paragraph para in comment.Paragraphs)
            {
                para.ParagraphFormat.StyleName = corporateStyleName;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CommentsStyled.docx");
        doc.Save(outputPath);
    }
}
