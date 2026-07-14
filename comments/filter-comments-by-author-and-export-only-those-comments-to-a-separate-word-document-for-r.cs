using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a source document with several comments from different authors.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment from Alice.
        srcBuilder.Writeln("Paragraph 1: Sample text.");
        Comment aliceComment = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        aliceComment.SetText("Review this paragraph, Alice.");
        srcBuilder.CurrentParagraph.AppendChild(aliceComment);

        // Second paragraph with a comment from Bob.
        srcBuilder.Writeln("Paragraph 2: Another sample.");
        Comment bobComment = new Comment(sourceDoc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Bob's suggestion.");
        srcBuilder.CurrentParagraph.AppendChild(bobComment);

        // Third paragraph with another comment from Alice.
        srcBuilder.Writeln("Paragraph 3: More text.");
        Comment aliceSecondComment = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        aliceSecondComment.SetText("Additional note from Alice.");
        srcBuilder.CurrentParagraph.AppendChild(aliceSecondComment);

        // Save the source document (optional, for inspection).
        sourceDoc.Save("SourceDocument.docx");

        // Enumerate all comments in the source document.
        var allComments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        // Filter comments authored by "Alice" (case‑insensitive).
        var filteredComments = allComments
                               .Where(c => string.Equals(c.Author, "Alice", StringComparison.OrdinalIgnoreCase))
                               .ToList();

        // Create a new document to hold the filtered comments.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Filtered Comments Report");
        reportBuilder.Writeln("------------------------");
        reportBuilder.Writeln();

        // Add each filtered comment's details to the report document.
        foreach (Comment comment in filteredComments)
        {
            reportBuilder.Writeln($"Author : {comment.Author}");
            reportBuilder.Writeln($"Date   : {comment.DateTime:yyyy-MM-dd HH:mm:ss}");
            reportBuilder.Writeln($"Text   : {comment.GetText().Trim()}");
            reportBuilder.Writeln(); // Blank line between entries.
        }

        // Save the report document containing only Alice's comments.
        reportDoc.Save("FilteredComments.docx");
    }
}
