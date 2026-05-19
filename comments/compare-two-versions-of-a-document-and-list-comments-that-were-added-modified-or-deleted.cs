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
        // Prepare a temporary folder for the sample files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create the original document with two comments.
        // -----------------------------------------------------------------
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        builder.Writeln("First paragraph in the original document.");

        // Comment #1
        Comment comment1 = new Comment(docOriginal, "Alice", "A", DateTime.Now);
        comment1.SetText("Original comment 1.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Comment #2
        Comment comment2 = new Comment(docOriginal, "Bob", "B", DateTime.Now);
        comment2.SetText("Original comment 2.");
        builder.CurrentParagraph.AppendChild(comment2);

        string originalPath = Path.Combine(outputDir, "Original.docx");
        docOriginal.Save(originalPath);

        // -----------------------------------------------------------------
        // Create the edited version: modify comment #1, delete comment #2,
        // and add a new comment #3.
        // -----------------------------------------------------------------
        Document docEdited = (Document)docOriginal.Clone(true);
        DocumentBuilder editBuilder = new DocumentBuilder(docEdited);

        // Modify comment #1 text.
        Comment? editedComment1 = docEdited.GetChildNodes(NodeType.Comment, true)
                                            .OfType<Comment>()
                                            .FirstOrDefault(c => c.Author == "Alice");
        if (editedComment1 != null)
        {
            editedComment1.SetText("Modified comment 1.");
        }

        // Delete comment #2.
        Comment? commentToDelete = docEdited.GetChildNodes(NodeType.Comment, true)
                                            .OfType<Comment>()
                                            .FirstOrDefault(c => c.Author == "Bob");
        commentToDelete?.Remove();

        // Add a new comment #3.
        editBuilder.Writeln("Additional paragraph in the edited document.");
        Comment comment3 = new Comment(docEdited, "Charlie", "C", DateTime.Now);
        comment3.SetText("New comment 3.");
        editBuilder.CurrentParagraph.AppendChild(comment3);

        string editedPath = Path.Combine(outputDir, "Edited.docx");
        docEdited.Save(editedPath);

        // -----------------------------------------------------------------
        // Compare the two documents and list comment changes.
        // -----------------------------------------------------------------
        // Load the documents (could also reuse the in‑memory objects).
        Document original = new Document(originalPath);
        Document edited = new Document(editedPath);

        // Gather comments from both documents.
        List<Comment> originalComments = original.GetChildNodes(NodeType.Comment, true)
                                                 .OfType<Comment>()
                                                 .ToList();

        List<Comment> editedComments = edited.GetChildNodes(NodeType.Comment, true)
                                             .OfType<Comment>()
                                             .ToList();

        // Index comments by their Id for quick lookup.
        var originalById = originalComments.ToDictionary(c => c.Id);
        var editedById = editedComments.ToDictionary(c => c.Id);

        // Detect added comments.
        var added = editedById.Keys.Except(originalById.Keys);
        foreach (int id in added)
        {
            Comment c = editedById[id];
            Console.WriteLine($"Added comment (Id={c.Id}): Author={c.Author}, Text=\"{c.GetText().Trim()}\"");
        }

        // Detect deleted comments.
        var deleted = originalById.Keys.Except(editedById.Keys);
        foreach (int id in deleted)
        {
            Comment c = originalById[id];
            Console.WriteLine($"Deleted comment (Id={c.Id}): Author={c.Author}, Text=\"{c.GetText().Trim()}\"");
        }

        // Detect modified comments (present in both versions but with different text).
        var possibleModified = originalById.Keys.Intersect(editedById.Keys);
        foreach (int id in possibleModified)
        {
            Comment originalComment = originalById[id];
            Comment editedComment = editedById[id];

            string originalText = originalComment.GetText().Trim();
            string editedText = editedComment.GetText().Trim();

            if (!string.Equals(originalText, editedText, StringComparison.Ordinal))
            {
                Console.WriteLine($"Modified comment (Id={id}): Author={editedComment.Author}");
                Console.WriteLine($"    Original: \"{originalText}\"");
                Console.WriteLine($"    Edited:   \"{editedText}\"");
            }
        }

        // -----------------------------------------------------------------
        // (Optional) Demonstrate the built‑in compare that produces revisions.
        // -----------------------------------------------------------------
        // The compare operation is not required for the listing logic above,
        // but it shows how Aspose.Words tracks comment changes as revisions.
        Document compareDoc = (Document)original.Clone(true);
        CompareOptions options = new CompareOptions(); // default: do not ignore comments
        compareDoc.Compare(edited, "Comparer", DateTime.Now, options);
        // The compareDoc now contains revisions for comment insertions/deletions/modifications.
        // For brevity we do not output the revisions here.
    }
}
