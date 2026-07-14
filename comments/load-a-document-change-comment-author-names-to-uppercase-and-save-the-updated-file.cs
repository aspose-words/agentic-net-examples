using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the input and output documents.
        string inputPath = "CommentsInput.docx";
        string outputPath = "CommentsOutput.docx";

        // Create a sample document that contains comments.
        CreateSampleDocument(inputPath);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve all comment nodes in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Convert each comment author's name to uppercase.
        foreach (Comment comment in comments)
        {
            if (!string.IsNullOrEmpty(comment.Author))
                comment.Author = comment.Author.ToUpperInvariant();
        }

        // Save the modified document.
        doc.Save(outputPath);
    }

    // Helper method to create a document with a couple of comments.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Create a comment for the first paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Review this paragraph.");
        // Attach the comment to the first paragraph.
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        firstPara.AppendChild(comment1);

        // Create a comment for the second paragraph.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Check the wording.");
        // Attach the comment to the second paragraph.
        Paragraph secondPara = doc.FirstSection.Body.Paragraphs[1];
        secondPara.AppendChild(comment2);

        // Save the document to the specified path.
        doc.Save(filePath);
    }
}
