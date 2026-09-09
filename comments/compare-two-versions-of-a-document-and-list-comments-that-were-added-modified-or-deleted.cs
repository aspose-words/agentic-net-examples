using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeWordsCommentsComparison
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for the sample files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create the original document with two comments.
            // -----------------------------------------------------------------
            Document originalDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(originalDoc);

            builder.Writeln("First paragraph.");

            Comment comment1 = new Comment(originalDoc, "Alice", "A", DateTime.Now);
            comment1.SetText("Original comment 1");
            builder.CurrentParagraph.AppendChild(comment1);

            builder.Writeln("Second paragraph.");

            Comment comment2 = new Comment(originalDoc, "Bob", "B", DateTime.Now);
            comment2.SetText("Original comment 2");
            builder.CurrentParagraph.AppendChild(comment2);

            string originalPath = Path.Combine(outputDir, "Original.docx");
            originalDoc.Save(originalPath);

            // -----------------------------------------------------------------
            // 2. Clone the original and modify it:
            //    - Change text of the first comment.
            //    - Delete the second comment.
            //    - Add a new third comment.
            // -----------------------------------------------------------------
            Document editedDoc = (Document)originalDoc.Clone(true);

            // Change text of the first comment.
            Comment editedComment1 = editedDoc.GetChildNodes(NodeType.Comment, true)
                                             .OfType<Comment>()
                                             .FirstOrDefault(c => c.Author == "Alice");
            if (editedComment1 != null && editedComment1.FirstParagraph?.Runs.Count > 0)
            {
                editedComment1.FirstParagraph.Runs[0].Text = "Modified comment 1";
            }

            // Delete the second comment.
            Comment editedComment2 = editedDoc.GetChildNodes(NodeType.Comment, true)
                                             .OfType<Comment>()
                                             .FirstOrDefault(c => c.Author == "Bob");
            editedComment2?.Remove();

            // Add a new third comment.
            DocumentBuilder editBuilder = new DocumentBuilder(editedDoc);
            editBuilder.Writeln("Third paragraph with a new comment.");

            Comment comment3 = new Comment(editedDoc, "Charlie", "C", DateTime.Now);
            comment3.SetText("New comment 3");
            editBuilder.CurrentParagraph.AppendChild(comment3);

            string editedPath = Path.Combine(outputDir, "Edited.docx");
            editedDoc.Save(editedPath);

            // -----------------------------------------------------------------
            // 3. Compare the two documents. The original document will receive revisions.
            // -----------------------------------------------------------------
            Document compareDoc = new Document(originalPath);
            Document compareTarget = new Document(editedPath);
            compareDoc.Compare(compareTarget, "Comparer", DateTime.Now);

            // -----------------------------------------------------------------
            // 4. Create a version of the original document with all revisions accepted.
            //    This represents the edited state.
            // -----------------------------------------------------------------
            Document finalDoc = (Document)compareDoc.Clone(true);
            finalDoc.Revisions.AcceptAll();

            // -----------------------------------------------------------------
            // 5. Enumerate comments in both the revision‑bearing document and the final document.
            // -----------------------------------------------------------------
            List<Comment> originalComments = compareDoc.GetChildNodes(NodeType.Comment, true)
                                                      .OfType<Comment>()
                                                      .ToList();

            List<Comment> finalComments = finalDoc.GetChildNodes(NodeType.Comment, true)
                                                  .OfType<Comment>()
                                                  .ToList();

            // Build dictionaries keyed by comment Id for quick lookup.
            Dictionary<int, Comment> originalById = originalComments.ToDictionary(c => c.Id);
            Dictionary<int, Comment> finalById = finalComments.ToDictionary(c => c.Id);

            // -----------------------------------------------------------------
            // 6. Determine added, deleted, and modified comments.
            // -----------------------------------------------------------------
            List<Comment> addedComments = finalComments.Where(c => !originalById.ContainsKey(c.Id)).ToList();
            List<Comment> deletedComments = originalComments.Where(c => !finalById.ContainsKey(c.Id)).ToList();

            List<(Comment Original, Comment Modified)> modifiedComments = new List<(Comment, Comment)>();
            foreach (var kvp in originalById)
            {
                int id = kvp.Key;
                Comment original = kvp.Value;
                if (finalById.TryGetValue(id, out Comment updated))
                {
                    string originalText = original.GetText().Trim();
                    string updatedText = updated.GetText().Trim();
                    if (!string.Equals(originalText, updatedText, StringComparison.Ordinal))
                    {
                        modifiedComments.Add((original, updated));
                    }
                }
            }

            // -----------------------------------------------------------------
            // 7. Output the results.
            // -----------------------------------------------------------------
            Console.WriteLine("=== Comment Comparison Report ===");
            Console.WriteLine();

            Console.WriteLine("Added Comments:");
            if (addedComments.Count == 0)
                Console.WriteLine("  (none)");
            else
                foreach (var c in addedComments)
                    Console.WriteLine($"  Author: {c.Author}, Text: \"{c.GetText().Trim()}\"");

            Console.WriteLine();

            Console.WriteLine("Deleted Comments:");
            if (deletedComments.Count == 0)
                Console.WriteLine("  (none)");
            else
                foreach (var c in deletedComments)
                    Console.WriteLine($"  Author: {c.Author}, Text: \"{c.GetText().Trim()}\"");

            Console.WriteLine();

            Console.WriteLine("Modified Comments:");
            if (modifiedComments.Count == 0)
                Console.WriteLine("  (none)");
            else
                foreach (var pair in modifiedComments)
                    Console.WriteLine($"  Author: {pair.Original.Author}, Original: \"{pair.Original.GetText().Trim()}\", Modified: \"{pair.Modified.GetText().Trim()}\"");

            // Save the documents for inspection (optional).
            compareDoc.Save(Path.Combine(outputDir, "ComparedWithRevisions.docx"));
            finalDoc.Save(Path.Combine(outputDir, "FinalAfterAccept.docx"));
        }
    }
}
