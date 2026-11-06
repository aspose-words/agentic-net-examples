using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create first version of the document with two comments.
        Document docV1 = new Document();
        DocumentBuilder builderV1 = new DocumentBuilder(docV1);

        // Paragraph 1 with a comment from Alice.
        builderV1.Writeln("First paragraph.");
        Comment comment1V1 = new Comment(docV1)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1V1.AppendChild(new Paragraph(docV1));
        comment1V1.FirstParagraph?.AppendChild(new Run(docV1, "First comment"));
        // Anchor the comment to the last paragraph.
        builderV1.CurrentParagraph?.AppendChild(comment1V1);

        // Paragraph 2 with a comment from Bob.
        builderV1.Writeln("Second paragraph.");
        Comment comment2V1 = new Comment(docV1)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now
        };
        comment2V1.AppendChild(new Paragraph(docV1));
        comment2V1.FirstParagraph?.AppendChild(new Run(docV1, "Second comment"));
        builderV1.CurrentParagraph?.AppendChild(comment2V1);

        // Save first version (optional, for inspection).
        docV1.Save("DocumentV1.docx");

        // Create second version of the document.
        Document docV2 = new Document();
        DocumentBuilder builderV2 = new DocumentBuilder(docV2);

        // Paragraph 1 with a modified comment from Alice.
        builderV2.Writeln("First paragraph.");
        Comment comment1V2 = new Comment(docV2)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1V2.AppendChild(new Paragraph(docV2));
        comment1V2.FirstParagraph?.AppendChild(new Run(docV2, "First comment modified"));
        builderV2.CurrentParagraph?.AppendChild(comment1V2);

        // Paragraph 2 without Bob's comment (deleted).

        builderV2.Writeln("Second paragraph.");

        // Paragraph 3 with a new comment from Charlie (added).
        builderV2.Writeln("Third paragraph.");
        Comment comment3V2 = new Comment(docV2)
        {
            Author = "Charlie",
            Initial = "C",
            DateTime = DateTime.Now
        };
        comment3V2.AppendChild(new Paragraph(docV2));
        comment3V2.FirstParagraph?.AppendChild(new Run(docV2, "New comment"));
        builderV2.CurrentParagraph?.AppendChild(comment3V2);

        // Save second version (optional, for inspection).
        docV2.Save("DocumentV2.docx");

        // Enumerate comments in both documents.
        var commentsV1 = docV1.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        var commentsV2 = docV2.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Build dictionaries keyed by author for simple comparison.
        var dictV1 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Comment c in commentsV1)
        {
            string author = c.Author ?? string.Empty;
            string text = c.GetText().Trim();
            dictV1[author] = text;
        }

        var dictV2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Comment c in commentsV2)
        {
            string author = c.Author ?? string.Empty;
            string text = c.GetText().Trim();
            dictV2[author] = text;
        }

        // Determine added comments.
        var added = dictV2.Keys.Except(dictV1.Keys, StringComparer.OrdinalIgnoreCase);
        // Determine deleted comments.
        var deleted = dictV1.Keys.Except(dictV2.Keys, StringComparer.OrdinalIgnoreCase);
        // Determine modified comments (same author, different text).
        var modified = dictV1.Keys.Intersect(dictV2.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(author => !string.Equals(dictV1[author], dictV2[author], StringComparison.Ordinal));

        // Output results.
        Console.WriteLine("Added Comments:");
        foreach (string author in added)
        {
            Console.WriteLine($"- Author: {author}, Text: {dictV2[author]}");
        }

        Console.WriteLine("\nDeleted Comments:");
        foreach (string author in deleted)
        {
            Console.WriteLine($"- Author: {author}, Text: {dictV1[author]}");
        }

        Console.WriteLine("\nModified Comments:");
        foreach (string author in modified)
        {
            Console.WriteLine($"- Author: {author}");
            Console.WriteLine($"  Old Text: {dictV1[author]}");
            Console.WriteLine($"  New Text: {dictV2[author]}");
        }
    }
}
