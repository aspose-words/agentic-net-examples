using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("First paragraph of the document.");
        builder.Writeln("Second paragraph of the document.");

        // Insert a comment on the first paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Initial review comment.");
        // Append the comment to the first paragraph.
        builder.MoveToDocumentStart();
        builder.CurrentParagraph.AppendChild(comment1);

        // Insert a second comment on the second paragraph.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Additional feedback needed.");
        // Append the comment to the last paragraph.
        builder.MoveToDocumentEnd();
        builder.CurrentParagraph.AppendChild(comment2);

        // Define a custom style that matches corporate branding for comment text.
        Style corporateCommentStyle = doc.Styles.Add(StyleType.Paragraph, "CorporateComment");
        corporateCommentStyle.Font.Name = "Arial";
        corporateCommentStyle.Font.Size = 10;
        corporateCommentStyle.Font.Color = Color.DarkBlue;
        corporateCommentStyle.Font.Bold = true;

        // Enumerate all comments in the document safely.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Apply the custom style to every paragraph inside each comment.
        foreach (Comment c in comments)
        {
            // Ensure the comment contains at least one paragraph.
            if (c.Paragraphs.Count == 0)
                c.EnsureMinimum();

            foreach (Paragraph p in c.Paragraphs)
            {
                p.ParagraphFormat.Style = corporateCommentStyle;
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "DocumentWithStyledComments.docx");
        doc.Save(outputPath);
    }
}
