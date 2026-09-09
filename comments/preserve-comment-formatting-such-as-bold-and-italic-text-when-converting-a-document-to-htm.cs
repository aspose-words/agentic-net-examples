using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This paragraph has a comment with formatted text.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Add a paragraph inside the comment to hold the formatted runs.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));

        // Bold text inside the comment.
        builder.Font.Bold = true;
        builder.Write("Bold text");

        // Italic text inside the comment.
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Write(" Italic text");

        // Reset formatting for any further content.
        builder.Font.Italic = false;

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the document to HTML. The comment's formatting is preserved in the HTML output.
        string htmlPath = Path.Combine(outputDir, "DocumentWithComment.html");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        doc.Save(htmlPath, saveOptions);
    }
}
