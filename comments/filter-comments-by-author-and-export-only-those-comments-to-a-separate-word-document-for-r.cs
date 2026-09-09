using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class FilterCommentsExample
{
    public static void Main()
    {
        // Create a source document with sample comments from different authors.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment from Alice.
        builder.Writeln("First paragraph.");
        AddComment(sourceDoc, builder.CurrentParagraph, "Alice", "AL", "First comment by Alice.");

        // Second paragraph with a comment from Bob.
        builder.Writeln("Second paragraph.");
        AddComment(sourceDoc, builder.CurrentParagraph, "Bob", "BO", "Comment by Bob.");

        // Third paragraph with another comment from Alice.
        builder.Writeln("Third paragraph.");
        AddComment(sourceDoc, builder.CurrentParagraph, "Alice", "AL", "Second comment by Alice.");

        // Save the source document (optional, for inspection).
        sourceDoc.Save("SourceDocument.docx");

        // Define the author whose comments we want to extract.
        const string targetAuthor = "Alice";

        // Enumerate all comments in the source document.
        var allComments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        // Filter comments by the specified author (case‑insensitive).
        var filteredComments = allComments
                               .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                               .ToList();

        // Create a new document that will contain the filtered comments.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln($"Comments authored by \"{targetAuthor}\":");
        reportBuilder.Writeln();

        // Recreate each filtered comment as plain text in the report document.
        foreach (Comment comment in filteredComments)
        {
            // Ensure the comment text is not null before trimming.
            string commentText = comment.GetText()?.Trim() ?? string.Empty;

            reportBuilder.Writeln($"Date: {comment.DateTime}");
            reportBuilder.Writeln($"Text: {commentText}");
            reportBuilder.Writeln(); // Add an empty line between comments.
        }

        // Save the report document containing only the filtered comments.
        reportDoc.Save("FilteredComments.docx");
    }

    // Helper method to add a simple comment to a paragraph.
    private static void AddComment(Document doc, Paragraph paragraph, string author, string initial, string text)
    {
        // Create a new comment with the specified metadata.
        Comment comment = new Comment(doc, author, initial, DateTime.Now);
        // Set the comment text; this creates at least one paragraph and run inside the comment.
        comment.SetText(text);

        // Append the comment to the paragraph.
        paragraph.AppendChild(comment);
    }
}
