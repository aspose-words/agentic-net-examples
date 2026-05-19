using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class DeleteCommentsByAuthor
{
    public static void Main()
    {
        // Prepare a temporary directory for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "CommentExample");
        Directory.CreateDirectory(workDir);

        // Paths for the input and output documents.
        string inputPath = Path.Combine(workDir, "input.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with several comments from different authors.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from John Doe.
        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment1.SetText("Comment by John Doe.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment from Jane Smith.
        builder.Writeln("Second paragraph.");
        Comment comment2 = new Comment(doc, "Jane Smith", "JS", DateTime.Now);
        comment2.SetText("Comment by Jane Smith.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Third paragraph with another comment from John Doe.
        builder.Writeln("Third paragraph.");
        Comment comment3 = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment3.SetText("Another comment by John Doe.");
        builder.CurrentParagraph.AppendChild(comment3);

        // Save the sample document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and delete all comments authored by "John Doe".
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);
        string targetAuthor = "John Doe";

        var commentsToDelete = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                                        .ToList();

        foreach (Comment c in commentsToDelete)
        {
            c.Remove();
        }

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Comments by \"{targetAuthor}\" have been removed.");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
