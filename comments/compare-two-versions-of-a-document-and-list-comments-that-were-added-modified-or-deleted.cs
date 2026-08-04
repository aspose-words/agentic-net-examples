using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace CommentComparisonDemo
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create the original document with three comments.
            Document originalDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(originalDoc);

            // First paragraph and comment.
            builder.Writeln("Paragraph 1.");
            Comment comment1 = new Comment(originalDoc, "Alice", "A", DateTime.Now);
            comment1.SetText("Original comment 1.");
            builder.CurrentParagraph.AppendChild(comment1);

            // Second paragraph and comment.
            builder.Writeln("Paragraph 2.");
            Comment comment2 = new Comment(originalDoc, "Bob", "B", DateTime.Now);
            comment2.SetText("Original comment 2.");
            builder.CurrentParagraph.AppendChild(comment2);

            // Third paragraph and comment.
            builder.Writeln("Paragraph 3.");
            Comment comment3 = new Comment(originalDoc, "Charlie", "C", DateTime.Now);
            comment3.SetText("Original comment 3.");
            builder.CurrentParagraph.AppendChild(comment3);

            // Save original document.
            string originalPath = Path.Combine(outputDir, "Original.docx");
            originalDoc.Save(originalPath);

            // Clone the original to create the edited version.
            Document editedDoc = (Document)originalDoc.Clone(true);

            // Locate comments in the edited document by their IDs.
            var editedComments = editedDoc.GetChildNodes(NodeType.Comment, true)
                                          .OfType<Comment>()
                                          .ToDictionary(c => c.Id);

            // Delete the second comment (Bob's comment).
            if (editedComments.TryGetValue(comment2.Id, out Comment? toDelete))
            {
                toDelete.Remove();
            }

            // Modify the text of the first comment (Alice's comment).
            if (editedComments.TryGetValue(comment1.Id, out Comment? toModify))
            {
                toModify.SetText("Modified comment 1.");
            }

            // Add a new comment (Dave's comment) to the last paragraph.
            Paragraph lastParagraph = editedDoc.FirstSection.Body.LastParagraph;
            Comment comment4 = new Comment(editedDoc, "Dave", "D", DateTime.Now);
            comment4.SetText("Newly added comment 4.");
            lastParagraph.AppendChild(comment4);

            // Save edited document.
            string editedPath = Path.Combine(outputDir, "Edited.docx");
            editedDoc.Save(editedPath);

            // -----------------------------------------------------------------
            // Compare the two documents and list comment differences.
            // -----------------------------------------------------------------

            // Perform a comparison; revisions will be added to the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

            // Gather comments from both versions.
            List<Comment> originalCommentList = originalDoc.GetChildNodes(NodeType.Comment, true)
                                                          .OfType<Comment>()
                                                          .ToList();

            List<Comment> editedCommentList = editedDoc.GetChildNodes(NodeType.Comment, true)
                                                        .OfType<Comment>()
                                                        .ToList();

            // Build dictionaries keyed by comment Id for quick lookup.
            var originalById = originalCommentList.ToDictionary(c => c.Id);
            var editedById = editedCommentList.ToDictionary(c => c.Id);

            // Track results.
            List<string> added = new List<string>();
            List<string> deleted = new List<string>();
            List<string> modified = new List<string>();

            // Detect added and modified comments.
            foreach (var editedPair in editedById)
            {
                int id = editedPair.Key;
                Comment editedComment = editedPair.Value;

                if (!originalById.ContainsKey(id))
                {
                    // New comment.
                    added.Add(FormatCommentInfo(editedComment));
                }
                else
                {
                    // Possible modification.
                    Comment originalComment = originalById[id];
                    string originalText = originalComment.GetText().Trim();
                    string editedText = editedComment.GetText().Trim();

                    if (!string.Equals(originalText, editedText, StringComparison.Ordinal))
                    {
                        modified.Add($"Id={id}, Author={editedComment.Author}, From=\"{originalText}\" To=\"{editedText}\"");
                    }
                }
            }

            // Detect deleted comments.
            foreach (var originalPair in originalById)
            {
                int id = originalPair.Key;
                if (!editedById.ContainsKey(id))
                {
                    deleted.Add(FormatCommentInfo(originalPair.Value));
                }
            }

            // Output the results.
            Console.WriteLine("Added comments:");
            foreach (string info in added) Console.WriteLine($"  {info}");

            Console.WriteLine("\nModified comments:");
            foreach (string info in modified) Console.WriteLine($"  {info}");

            Console.WriteLine("\nDeleted comments:");
            foreach (string info in deleted) Console.WriteLine($"  {info}");
        }

        // Helper to format comment information for display.
        private static string FormatCommentInfo(Comment comment)
        {
            string text = comment.GetText().Trim();
            return $"Id={comment.Id}, Author={comment.Author}, Text=\"{text}\"";
        }
    }
}
