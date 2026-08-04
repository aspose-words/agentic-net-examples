using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class DeleteCommentsByAuthor
{
    public static void Main()
    {
        // Define the author whose comments will be removed.
        const string targetAuthor = "John Doe";

        // Create a sample document with comments from different authors.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("This is the first paragraph.");

        // Add a comment authored by John Doe.
        AddComment(builder.CurrentParagraph, doc, "John Doe", "JD", "Comment from John.");

        // Add a second paragraph.
        builder.Writeln("This is the second paragraph.");

        // Add a comment authored by Jane Smith.
        AddComment(builder.CurrentParagraph, doc, "Jane Smith", "JS", "Comment from Jane.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // Load the document (simulating a separate processing step).
        Document loadedDoc = new Document(originalPath);

        // Find all comments authored by the target author.
        var commentsToRemove = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                                        .ToList();

        // Remove each matching comment.
        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the cleaned document.
        string cleanedPath = Path.Combine(outputDir, "cleaned.docx");
        loadedDoc.Save(cleanedPath);
    }

    // Helper method to add a comment anchored to a paragraph.
    private static void AddComment(Paragraph paragraph, Document doc, string author, string initial, string commentText)
    {
        // Create a new comment.
        Comment comment = new Comment(doc, author, initial, DateTime.Now);
        comment.SetText(commentText);

        // Insert the comment range start, the commented text, the range end, and the comment node.
        CommentRangeStart rangeStart = new CommentRangeStart(doc, comment.Id);
        CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, comment.Id);
        Run commentedRun = new Run(doc, "Commented text.");

        paragraph.AppendChild(rangeStart);
        paragraph.AppendChild(commentedRun);
        paragraph.AppendChild(rangeEnd);
        paragraph.AppendChild(comment);
    }
}
