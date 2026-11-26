using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph with a comment.");
        builder.Writeln("Second paragraph without a comment.");

        // Create a comment for the first paragraph.
        Comment comment1 = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph for accuracy."));
        // Attach the comment to the first paragraph.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        firstParagraph?.AppendChild(comment1);

        // Create a comment for the second paragraph.
        Comment comment2 = new Comment(doc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        // Attach the comment to the second paragraph.
        Paragraph? secondParagraph = firstParagraph?.NextSibling as Paragraph;
        secondParagraph?.AppendChild(comment2);

        // Define a custom character style for corporate branding.
        const string styleName = "CorporateComment";
        Style? existingStyle = doc.Styles[styleName];
        if (existingStyle == null)
        {
            Style corporateStyle = doc.Styles.Add(StyleType.Character, styleName);
            corporateStyle.Font.Color = Color.Blue;
            corporateStyle.Font.Size = 10;
            corporateStyle.Font.Bold = true;
            corporateStyle.Font.Italic = false;
        }

        // Apply the custom style to all comment text runs.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            var runs = comment.GetChildNodes(NodeType.Run, true)
                              .OfType<Run>()
                              .ToList();

            foreach (Run run in runs)
            {
                run.Font.StyleName = styleName;
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "styled-comments.docx");
        doc.Save(outputPath);
    }
}
