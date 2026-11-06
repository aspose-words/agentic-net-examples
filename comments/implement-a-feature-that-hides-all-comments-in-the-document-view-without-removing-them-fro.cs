using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class HideCommentsExample
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment attached.");

        // Create a comment with visible text.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add a paragraph and a run inside the comment so it contains text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the current paragraph (the one we just wrote).
        Paragraph? targetParagraph = builder.CurrentParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document with visible comments.
        string originalPath = Path.Combine(outputDir, "original-with-comments.docx");
        doc.Save(originalPath);

        // NOTE: Aspose.Words does not provide a direct API to hide comments in a DOCX file
        // while keeping them in the document. The typical approach is to control comment
        // visibility in the Word UI (e.g., via Word's "Show Comments" option). Therefore,
        // we simply save a copy of the same document to demonstrate that comments are
        // still present and can be inspected programmatically.

        string hiddenPath = Path.Combine(outputDir, "hidden-comments.docx");
        doc.Save(hiddenPath);

        // Reload the saved document to verify comments still exist.
        Document reloaded = new Document(hiddenPath);
        var comments = reloaded.GetChildNodes(NodeType.Comment, true)
                               .OfType<Comment>()
                               .ToList();

        // Output verification information.
        Console.WriteLine($"Total comments in the saved document: {comments.Count}");
        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author}: {c.GetText().Trim()}");
        }
    }
}
