using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string sourcePath = "Source.docx";
        const string reportPath = "FilteredComments.docx";

        // ------------------------------------------------------------
        // 1. Create a sample document that contains comments from two authors.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment from Alice.
        builder.Writeln("First paragraph.");
        Comment commentAlice = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        commentAlice.SetText("Alice's comment on the first paragraph.");
        builder.CurrentParagraph.AppendChild(commentAlice);

        // Second paragraph with a comment from Bob.
        builder.Writeln("Second paragraph.");
        Comment commentBob = new Comment(sourceDoc, "Bob", "B", DateTime.Now.AddMinutes(-5));
        commentBob.SetText("Bob's comment on the second paragraph.");
        builder.CurrentParagraph.AppendChild(commentBob);

        // Save the source document to disk.
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Load the document and enumerate all comments.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        var allComments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList(); // Safe copy for further processing.

        // ------------------------------------------------------------
        // 3. Filter comments by the desired author.
        // ------------------------------------------------------------
        const string targetAuthor = "Alice";

        var filteredComments = allComments
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // ------------------------------------------------------------
        // 4. Create a new document that will contain only the filtered comments.
        // ------------------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln($"Comments authored by \"{targetAuthor}\":");
        reportBuilder.Writeln(); // Blank line.

        foreach (Comment c in filteredComments)
        {
            // Write comment metadata and text to the report.
            reportBuilder.Writeln($"Author: {c.Author}");
            reportBuilder.Writeln($"Date: {c.DateTime:yyyy-MM-dd HH:mm}");
            reportBuilder.Writeln($"Text: {c.GetText().Trim()}");
            reportBuilder.Writeln(); // Separate entries.
        }

        // ------------------------------------------------------------
        // 5. Save the report document.
        // ------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
