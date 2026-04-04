using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class DeleteCommentsByAuthor
{
    public static void Main()
    {
        // Define file names.
        const string sourceFile = "sample.docx";
        const string resultFile = "output.docx";

        // Create a sample document with comments.
        CreateSampleDocument(sourceFile);

        // Load the document from the file.
        Document doc = new Document(sourceFile);

        // Author whose comments should be removed.
        const string targetAuthor = "John Doe";

        // Find all comments authored by the target author.
        var commentsToDelete = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                                  .ToList();

        // Remove each matching comment safely.
        foreach (Comment comment in commentsToDelete)
        {
            comment.Remove();
        }

        // Save the modified document.
        doc.Save(resultFile);
    }

    // Creates a simple document with two comments, one by John Doe and another by Jane Smith.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("This is the first paragraph.");

        // Add a comment by John Doe.
        Comment commentJohn = new Comment(doc, "John Doe", "JD", DateTime.Now);
        commentJohn.SetText("Comment from John.");
        builder.CurrentParagraph.AppendChild(commentJohn);

        // Second paragraph.
        builder.Writeln("This is the second paragraph.");

        // Add a comment by Jane Smith.
        Comment commentJane = new Comment(doc, "Jane Smith", "JS", DateTime.Now);
        commentJane.SetText("Comment from Jane.");
        builder.CurrentParagraph.AppendChild(commentJane);

        // Save the sample document.
        doc.Save(filePath);
    }
}
