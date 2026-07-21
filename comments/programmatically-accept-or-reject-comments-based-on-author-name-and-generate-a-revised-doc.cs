using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 with a comment from Alice.
        builder.Write("Paragraph 1: This is the first paragraph.");
        Comment aliceComment = new Comment(doc, "Alice", "A", DateTime.Now);
        aliceComment.SetText("Alice's review comment.");
        builder.CurrentParagraph?.AppendChild(aliceComment);
        builder.Writeln(); // Finish the paragraph.

        // Paragraph 2 with a comment from Bob.
        builder.Write("Paragraph 2: This is the second paragraph.");
        Comment bobComment = new Comment(doc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Bob's review comment.");
        builder.CurrentParagraph?.AppendChild(bobComment);
        builder.Writeln();

        // Save the original document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // Filter comments: keep only those authored by "Alice".
        var commentsToRemove = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .Where(c => !string.Equals(c.Author, "Alice", StringComparison.OrdinalIgnoreCase))
                                  .ToList();

        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the revised document.
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        doc.Save(revisedPath);

        // Output file locations.
        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Revised document saved to: {revisedPath}");
    }
}
