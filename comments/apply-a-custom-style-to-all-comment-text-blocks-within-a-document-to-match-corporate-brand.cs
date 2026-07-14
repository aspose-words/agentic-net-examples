using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample paragraphs.
        builder.Writeln("This is the first paragraph of the document.");
        builder.Writeln("This is the second paragraph, which will have a comment.");

        // Create a comment anchored to the current paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Please review this paragraph for accuracy.");
        // Append the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment1);

        // Add another paragraph with a comment.
        builder.Writeln("Another paragraph that also needs attention.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing this sentence.");
        builder.CurrentParagraph.AppendChild(comment2);

        // -----------------------------------------------------------------
        // Define a custom character style that matches corporate branding.
        // -----------------------------------------------------------------
        const string corporateStyleName = "CorporateComment";
        // Create the style if it does not already exist.
        Style corporateStyle = doc.Styles[corporateStyleName];
        if (corporateStyle == null)
        {
            corporateStyle = doc.Styles.Add(StyleType.Character, corporateStyleName);
            corporateStyle.Font.Name = "Arial";
            corporateStyle.Font.Size = 10;
            corporateStyle.Font.Color = Color.DarkBlue;
            corporateStyle.Font.Bold = true;
            corporateStyle.Font.Italic = false;
        }

        // ---------------------------------------------------------------
        // Apply the custom style to all comment text blocks in the document.
        // ---------------------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // Each comment may contain multiple paragraphs.
            foreach (Paragraph para in c.Paragraphs)
            {
                // Apply the style to every run within the paragraph.
                foreach (Run run in para.Runs)
                {
                    run.Font.StyleName = corporateStyleName;
                }
            }
        }

        // Save the modified document to the working directory.
        doc.Save("DocumentWithStyledComments.docx");
    }
}
