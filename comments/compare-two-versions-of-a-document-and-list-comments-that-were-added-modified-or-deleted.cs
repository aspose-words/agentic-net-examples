using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -------------------- Create original document with comments --------------------
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(docOriginal, "Alice", "A", DateTime.Now);
        comment1.SetText("Original comment 1.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(docOriginal, "Bob", "B", DateTime.Now);
        comment2.SetText("Original comment 2.");
        builder.CurrentParagraph.AppendChild(comment2);

        string originalPath = Path.Combine(outputDir, "original.docx");
        docOriginal.Save(originalPath);

        // -------------------- Create edited document (clone and modify) --------------------
        Document docEdited = (Document)docOriginal.Clone(true);
        DocumentBuilder editedBuilder = new DocumentBuilder(docEdited);

        // Remove the first comment (Alice's comment).
        var originalCommentsInEdited = docEdited.GetChildNodes(NodeType.Comment, true)
                                                .OfType<Comment>()
                                                .ToList();
        Comment? commentToRemove = originalCommentsInEdited.FirstOrDefault(c => c.Author == "Alice");
        commentToRemove?.Remove();

        // Modify Bob's comment text.
        Comment? commentToModify = docEdited.GetChildNodes(NodeType.Comment, true)
                                            .OfType<Comment>()
                                            .FirstOrDefault(c => c.Author == "Bob");
        if (commentToModify != null)
        {
            commentToModify.SetText("Edited comment 2.");
        }

        // Add a new comment on a new paragraph.
        editedBuilder.Writeln("This is the third paragraph.");
        Comment newComment = new Comment(docEdited, "Charlie", "C", DateTime.Now);
        newComment.SetText("Newly added comment.");
        editedBuilder.CurrentParagraph.AppendChild(newComment);

        string editedPath = Path.Combine(outputDir, "edited.docx");
        docEdited.Save(editedPath);

        // -------------------- Compare documents (optional, to generate revisions) --------------------
        Document docForComparison = (Document)docOriginal.Clone(true);
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreComments = false // Ensure comments are taken into account.
        };
        docForComparison.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);
        // Revisions are generated in docForComparison, but we will report differences manually.

        // -------------------- Analyze comment differences --------------------
        var originalComments = docOriginal.GetChildNodes(NodeType.Comment, true)
                                          .OfType<Comment>()
                                          .ToList();
        var editedComments = docEdited.GetChildNodes(NodeType.Comment, true)
                                      .OfType<Comment>()
                                      .ToList();

        var originalById = originalComments.ToDictionary(c => c.Id);
        var editedById = editedComments.ToDictionary(c => c.Id);

        // Added comments: present in edited but not in original.
        var added = editedById.Keys.Except(originalById.Keys)
                                   .Select(id => editedById[id])
                                   .ToList();

        // Deleted comments: present in original but not in edited.
        var deleted = originalById.Keys.Except(editedById.Keys)
                                      .Select(id => originalById[id])
                                      .ToList();

        // Modified comments: same Id in both, but text differs.
        var modified = originalById.Keys.Intersect(editedById.Keys)
                                        .Where(id => !string.Equals(
                                            originalById[id].GetText().Trim(),
                                            editedById[id].GetText().Trim(),
                                            StringComparison.Ordinal))
                                        .Select(id => new
                                        {
                                            Original = originalById[id],
                                            Edited = editedById[id]
                                        })
                                        .ToList();

        // -------------------- Output results --------------------
        Console.WriteLine("=== Comment Comparison Report ===");
        Console.WriteLine();

        Console.WriteLine("Added Comments:");
        if (added.Count == 0)
            Console.WriteLine("  (none)");
        else
        {
            foreach (var c in added)
            {
                Console.WriteLine($"  Author: {c.Author}, Text: {c.GetText().Trim()}");
            }
        }

        Console.WriteLine();
        Console.WriteLine("Deleted Comments:");
        if (deleted.Count == 0)
            Console.WriteLine("  (none)");
        else
        {
            foreach (var c in deleted)
            {
                Console.WriteLine($"  Author: {c.Author}, Text: {c.GetText().Trim()}");
            }
        }

        Console.WriteLine();
        Console.WriteLine("Modified Comments:");
        if (modified.Count == 0)
            Console.WriteLine("  (none)");
        else
        {
            foreach (var pair in modified)
            {
                Console.WriteLine($"  Original - Author: {pair.Original.Author}, Text: {pair.Original.GetText().Trim()}");
                Console.WriteLine($"  Edited   - Author: {pair.Edited.Author}, Text: {pair.Edited.GetText().Trim()}");
                Console.WriteLine();
            }
        }
    }
}
