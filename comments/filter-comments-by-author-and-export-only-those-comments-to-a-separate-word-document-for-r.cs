using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class FilterCommentsExample
{
    public static void Main()
    {
        // Create a source document with sample paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment by Alice.
        builder.Writeln("This is the first paragraph.");
        Comment aliceComment = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        aliceComment.SetText("Alice's review comment.");
        builder.CurrentParagraph.AppendChild(aliceComment);

        // Second paragraph with a comment by Bob.
        builder.Writeln("This is the second paragraph.");
        Comment bobComment = new Comment(sourceDoc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Bob's feedback.");
        builder.CurrentParagraph.AppendChild(bobComment);

        // Third paragraph with another comment by Alice.
        builder.Writeln("This is the third paragraph.");
        Comment aliceSecondComment = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        aliceSecondComment.SetText("Another note from Alice.");
        builder.CurrentParagraph.AppendChild(aliceSecondComment);

        // Define the author whose comments we want to extract.
        const string targetAuthor = "Alice";

        // Enumerate all comments in the source document and filter by author.
        var filteredComments = sourceDoc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create a new document to hold the filtered comments.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln($"Comments authored by \"{targetAuthor}\":");
        reportBuilder.Writeln();

        // Write each filtered comment's details into the report document.
        foreach (Comment comment in filteredComments)
        {
            reportBuilder.Writeln($"Author : {comment.Author}");
            reportBuilder.Writeln($"Date   : {comment.DateTime:O}");
            reportBuilder.Writeln($"Text   : {comment.GetText().Trim()}");
            reportBuilder.Writeln(); // Add an empty line between comments.
        }

        // Save the report document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FilteredComments.docx");
        reportDoc.Save(outputPath);
    }
}
