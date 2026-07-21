using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string sourceFile = "SourceDocument.docx";
        const string reportFile = "FilteredComments.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample source document with comments.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment from Alice.
        srcBuilder.Writeln("First paragraph with a comment.");
        Comment aliceComment = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        aliceComment.SetText("This is Alice's comment.");
        srcBuilder.CurrentParagraph.AppendChild(aliceComment);

        // Second paragraph with a comment from Bob.
        srcBuilder.Writeln("Second paragraph with a comment.");
        Comment bobComment = new Comment(sourceDoc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Bob's remark goes here.");
        srcBuilder.CurrentParagraph.AppendChild(bobComment);

        // Third paragraph with another comment from Alice.
        srcBuilder.Writeln("Third paragraph, another comment.");
        Comment aliceSecond = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        aliceSecond.SetText("Alice adds a second comment.");
        srcBuilder.CurrentParagraph.AppendChild(aliceSecond);

        // Save the source document.
        sourceDoc.Save(sourceFile);

        // -----------------------------------------------------------------
        // Step 2: Load the source document and filter comments by author.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceFile);
        const string targetAuthor = "Alice";

        var filteredComments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // -----------------------------------------------------------------
        // Step 3: Create a new document that will contain only the filtered comments.
        // -----------------------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        if (filteredComments.Count == 0)
        {
            reportBuilder.Writeln($"No comments found for author \"{targetAuthor}\".");
        }
        else
        {
            foreach (Comment comment in filteredComments)
            {
                // Write comment metadata and text.
                reportBuilder.Writeln($"Author: {comment.Author}");
                reportBuilder.Writeln($"Date: {comment.DateTime:O}");
                reportBuilder.Writeln($"Text: {comment.GetText().Trim()}");
                reportBuilder.Writeln(); // Blank line between comments.
            }
        }

        // Save the report document.
        reportDoc.Save(reportFile);
    }
}
