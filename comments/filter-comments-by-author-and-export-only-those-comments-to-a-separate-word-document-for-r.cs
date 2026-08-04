using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the working directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "source.docx");
        string exportPath = Path.Combine(Directory.GetCurrentDirectory(), "filtered-comments.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with several comments from different authors.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment from Alice.
        builder.Writeln("First paragraph.");
        Comment commentAlice = new Comment(sourceDoc, "Alice", "A", DateTime.Now);
        commentAlice.SetText("Review the first paragraph.");
        builder.CurrentParagraph.AppendChild(commentAlice);

        // Second paragraph with a comment from Bob.
        builder.Writeln("Second paragraph.");
        Comment commentBob = new Comment(sourceDoc, "Bob", "B", DateTime.Now.AddMinutes(-5));
        commentBob.SetText("Check the data in this paragraph.");
        builder.CurrentParagraph.AppendChild(commentBob);

        // Third paragraph with another comment from Alice.
        builder.Writeln("Third paragraph.");
        Comment commentAlice2 = new Comment(sourceDoc, "Alice", "A", DateTime.Now.AddHours(-1));
        commentAlice2.SetText("Consider rephrasing this sentence.");
        builder.CurrentParagraph.AppendChild(commentAlice2);

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and filter comments by author.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        const string targetAuthor = "Alice";

        var filteredComments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // -----------------------------------------------------------------
        // 3. Create a new document that will contain only the filtered comments.
        // -----------------------------------------------------------------
        Document exportDoc = new Document();
        DocumentBuilder exportBuilder = new DocumentBuilder(exportDoc);

        exportBuilder.Writeln($"Comments authored by \"{targetAuthor}\":");
        exportBuilder.Writeln();

        foreach (Comment c in filteredComments)
        {
            exportBuilder.Writeln($"Author : {c.Author}");
            exportBuilder.Writeln($"Date   : {c.DateTime:yyyy-MM-dd HH:mm}");
            exportBuilder.Writeln($"Text   : {c.GetText().Trim()}");
            exportBuilder.Writeln(); // Blank line between comments.
        }

        // Save the filtered comments document.
        exportDoc.Save(exportPath);
    }
}
