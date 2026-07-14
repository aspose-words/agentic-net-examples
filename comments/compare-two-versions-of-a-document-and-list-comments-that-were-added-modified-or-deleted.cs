using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create the original document ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // First paragraph with a comment.
        builder.Writeln("Paragraph 1.");
        Comment comment1 = new Comment(original, "Alice", "A", DateTime.Now);
        comment1.AppendChild(new Paragraph(original));
        comment1.FirstParagraph.AppendChild(new Run(original, "Original comment 1."));
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with another comment.
        builder.Writeln("Paragraph 2.");
        Comment comment2 = new Comment(original, "Charlie", "C", DateTime.Now);
        comment2.AppendChild(new Paragraph(original));
        comment2.FirstParagraph.AppendChild(new Run(original, "Original comment 2."));
        builder.CurrentParagraph.AppendChild(comment2);

        original.Save(Path.Combine(outputDir, "Original.docx"));

        // ---------- Clone and modify to create the edited document ----------
        Document edited = (Document)original.Clone(true);
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);

        // Capture comment IDs for later reference.
        int comment1Id = comment1.Id;
        int comment2Id = comment2.Id;

        // Modify the text of the first comment.
        Comment? editedComment1 = edited.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .FirstOrDefault(c => c.Id == comment1Id);
        if (editedComment1 != null)
        {
            // Ensure the comment has a paragraph and run before changing text.
            if (editedComment1.FirstParagraph == null)
                editedComment1.AppendChild(new Paragraph(edited));
            if (!editedComment1.FirstParagraph.Runs.Any())
                editedComment1.FirstParagraph.AppendChild(new Run(edited, string.Empty));

            editedComment1.FirstParagraph.Runs[0].Text = "Modified comment 1.";
        }

        // Delete the second comment.
        Comment? editedComment2 = edited.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .FirstOrDefault(c => c.Id == comment2Id);
        editedComment2?.Remove();

        // Add a new comment.
        editedBuilder.Writeln("Paragraph 3.");
        Comment comment3 = new Comment(edited, "Bob", "B", DateTime.Now);
        comment3.AppendChild(new Paragraph(edited));
        comment3.FirstParagraph.AppendChild(new Run(edited, "New comment added in edited version."));
        editedBuilder.CurrentParagraph.AppendChild(comment3);

        edited.Save(Path.Combine(outputDir, "Edited.docx"));

        // ---------- Capture comment collections before comparison ----------
        List<Comment> originalComments = original.GetChildNodes(NodeType.Comment, true)
                                                 .OfType<Comment>()
                                                 .ToList();

        List<Comment> editedComments = edited.GetChildNodes(NodeType.Comment, true)
                                             .OfType<Comment>()
                                             .ToList();

        // ---------- Compare documents ----------
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreComments = false, // Track comment changes.
            Target = ComparisonTargetType.New
        };

        // Perform comparison (original will contain revisions after this call).
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);
        original.Save(Path.Combine(outputDir, "Compared.docx"));

        // ---------- Determine added, deleted, and modified comments ----------
        var originalDict = originalComments.ToDictionary(c => c.Id);
        var editedDict = editedComments.ToDictionary(c => c.Id);

        // Added comments: present only in edited.
        List<Comment> added = editedDict.Values
                                         .Where(c => !originalDict.ContainsKey(c.Id))
                                         .ToList();

        // Deleted comments: present only in original.
        List<Comment> deleted = originalDict.Values
                                            .Where(c => !editedDict.ContainsKey(c.Id))
                                            .ToList();

        // Modified comments: same Id exists in both but text differs.
        List<(Comment Original, Comment Edited)> modified = originalDict.Values
            .Where(c => editedDict.ContainsKey(c.Id))
            .Select(c => (Original: c, Edited: editedDict[c.Id]))
            .Where(pair => !string.Equals(pair.Original.GetText().Trim(),
                                          pair.Edited.GetText().Trim(),
                                          StringComparison.Ordinal))
            .ToList();

        // ---------- Output results ----------
        Console.WriteLine("Added comments:");
        foreach (Comment c in added)
            Console.WriteLine($"- Id:{c.Id} Author:{c.Author} Text:{c.GetText().Trim()}");

        Console.WriteLine("\nDeleted comments:");
        foreach (Comment c in deleted)
            Console.WriteLine($"- Id:{c.Id} Author:{c.Author} Text:{c.GetText().Trim()}");

        Console.WriteLine("\nModified comments:");
        foreach (var pair in modified)
            Console.WriteLine($"- Id:{pair.Original.Id} Author:{pair.Original.Author}\n  Original: {pair.Original.GetText().Trim()}\n  Edited:   {pair.Edited.GetText().Trim()}");
    }
}
