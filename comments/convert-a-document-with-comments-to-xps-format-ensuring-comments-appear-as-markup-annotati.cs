using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph that will have a comment.");
        AddCommentToCurrentParagraph(doc, builder, "John Doe", "JD", "First comment text.");

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph with another comment.");
        AddCommentToCurrentParagraph(doc, builder, "Jane Smith", "JS", "Second comment text.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the document as XPS. Comments are preserved as markup annotations.
        string xpsPath = Path.Combine(outputDir, "DocumentWithComments.xps");
        doc.Save(xpsPath, SaveFormat.Xps);

        // Enumerate comments and write their details to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            string text = comment.GetText().Trim();
            Console.WriteLine($"Author: {comment.Author}, Initials: {comment.Initial}, Date: {comment.DateTime:u}");
            Console.WriteLine($"Comment Text: {text}");
            Console.WriteLine();
        }

        // Indicate completion.
        Console.WriteLine($"Document saved to: {xpsPath}");
    }

    private static void AddCommentToCurrentParagraph(Document doc, DocumentBuilder builder, string author, string initial, string commentText)
    {
        // Safely get the paragraph that the builder is currently on.
        Paragraph? paragraph = builder.CurrentParagraph;
        if (paragraph == null)
            return;

        // Create a new comment node.
        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initial,
            DateTime = DateTime.Now
        };

        // Add a paragraph and run inside the comment to hold the comment text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, commentText));
        comment.AppendChild(commentParagraph);

        // Append the comment to the current paragraph.
        paragraph.AppendChild(comment);
    }
}
