using System;
using System.IO;
using System.Linq;
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
        builder.Writeln("This paragraph will have a comment attached to it.");

        // Create a comment, set its metadata and text.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review the wording of this paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Add a second comment to demonstrate multiple comments.
        builder.Writeln("Another paragraph without a comment.");
        builder.Writeln("Paragraph with a second comment.");
        Comment secondComment = new Comment(doc, "Bob", "B", DateTime.Now);
        secondComment.SetText("Consider shortening this sentence.");
        builder.CurrentParagraph.AppendChild(secondComment);

        // Optional: enumerate comments and write their details to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Date: {c.DateTime}, Text: {c.GetText().Trim()}");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document to XPS format. Comments will be rendered as markup annotations.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        string xpsPath = Path.Combine(outputDir, "DocumentWithComments.xps");
        doc.Save(xpsPath, xpsOptions);

        Console.WriteLine($"Document saved to XPS at: {xpsPath}");
    }
}
