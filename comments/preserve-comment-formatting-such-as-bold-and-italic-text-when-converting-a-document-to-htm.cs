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

        // Add a paragraph that will contain a comment.
        builder.Writeln("This paragraph will have a comment with formatted text.");

        // Create a comment anchored to the current paragraph.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);

        // Build the comment body with bold and italic runs.
        Paragraph commentParagraph = new Paragraph(doc);

        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;

        Run italicRun = new Run(doc, " and italic text");
        italicRun.Font.Italic = true;

        commentParagraph.AppendChild(boldRun);
        commentParagraph.AppendChild(italicRun);
        comment.AppendChild(commentParagraph);

        // Attach the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the document in DOCX format (optional, for inspection).
        string docxPath = Path.Combine(outputDir, "CommentFormatting.docx");
        doc.Save(docxPath, SaveFormat.Docx);

        // Save the document as HTML. The comment's formatted text will be preserved in the HTML output.
        string htmlPath = Path.Combine(outputDir, "CommentFormatting.html");
        doc.Save(htmlPath, SaveFormat.Html);

        // Enumerate comments and write their plain text to the console (verification).
        var comments = doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}");
            Console.WriteLine($"Text: {c.GetText().Trim()}");
        }
    }
}
