using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and add some content with comments
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph without comment
        builder.Writeln("First paragraph without comment.");

        // Paragraph that will have a comment attached
        builder.Writeln("Paragraph that will have a comment attached.");
        Paragraph para1 = doc.FirstSection.Body.LastParagraph ?? throw new InvalidOperationException("Paragraph not found.");

        // Create first comment
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice Johnson",
            Initial = "AJ",
            DateTime = DateTime.Now
        };
        Paragraph commentPara1 = new Paragraph(doc);
        commentPara1.AppendChild(new Run(doc, "Please review this paragraph for accuracy."));
        comment1.AppendChild(commentPara1);
        para1.AppendChild(comment1);

        // Another paragraph needing attention
        builder.Writeln("Another paragraph needing attention.");
        Paragraph para2 = doc.FirstSection.Body.LastParagraph ?? throw new InvalidOperationException("Paragraph not found.");

        // Create second comment
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob Smith",
            Initial = "BS",
            DateTime = DateTime.Now.AddDays(-1)
        };
        Paragraph commentPara2 = new Paragraph(doc);
        commentPara2.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        comment2.AppendChild(commentPara2);
        para2.AppendChild(comment2);

        // Define a custom style for comment text to match corporate branding
        Style corporateStyle = doc.Styles.Add(StyleType.Paragraph, "CorporateComment");
        corporateStyle.Font.Color = Color.Blue;
        corporateStyle.Font.Bold = true;
        corporateStyle.Font.Italic = true;
        corporateStyle.Font.Size = 10;

        // Apply the custom style to all comment text blocks
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            var commentParagraphs = c.GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>()
                .ToList();

            foreach (Paragraph p in commentParagraphs)
            {
                p.ParagraphFormat.Style = corporateStyle;
            }
        }

        // Save the resulting document
        string outputPath = Path.Combine(outputDir, "styled-comments.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine($"Document with styled comments saved to: {outputPath}");
    }
}
