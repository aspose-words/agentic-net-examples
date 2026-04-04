using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    // Simple DTO to hold comment data for comparison.
    private class CommentInfo
    {
        public int Id { get; set; }
        public string Author { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the original document with two comments.
        // -----------------------------------------------------------------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(original, "Alice", "A", DateTime.Now);
        comment1.SetText("Original comment 1.");
        builder.CurrentParagraph.AppendChild(comment1);

        builder.Writeln("Second paragraph.");
        Comment comment2 = new Comment(original, "Bob", "B", DateTime.Now);
        comment2.SetText("Original comment 2.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the original document (optional, just for inspection).
        original.Save("Original.docx");

        // -----------------------------------------------------------------
        // 2. Clone the original and modify it:
        //    - Change text of the first comment.
        //    - Delete the second comment.
        //    - Add a new third comment.
        // -----------------------------------------------------------------
        Document edited = (Document)original.Clone(true);
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);

        // Change text of the first comment.
        Comment? editedComment1 = edited.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .FirstOrDefault(c => c.Id == comment1.Id);
        editedComment1?.SetText("Edited comment 1.");

        // Delete the second comment.
        Comment? editedComment2 = edited.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .FirstOrDefault(c => c.Id == comment2.Id);
        editedComment2?.Remove();

        // Add a new comment.
        editedBuilder.Writeln("Third paragraph.");
        Comment comment3 = new Comment(edited, "Charlie", "C", DateTime.Now);
        comment3.SetText("New comment 3.");
        editedBuilder.CurrentParagraph.AppendChild(comment3);

        // Save the edited document (optional).
        edited.Save("Edited.docx");

        // -----------------------------------------------------------------
        // 3. Extract comment information from both documents.
        // -----------------------------------------------------------------
        List<CommentInfo> originalComments = GetCommentsInfo(original);
        List<CommentInfo> editedComments   = GetCommentsInfo(edited);

        // Build dictionaries keyed by comment Id for fast lookup.
        var originalDict = originalComments.ToDictionary(c => c.Id);
        var editedDict   = editedComments.ToDictionary(c => c.Id);

        // -----------------------------------------------------------------
        // 4. Determine added, deleted, and modified comments.
        // -----------------------------------------------------------------
        var addedComments = editedDict.Keys.Except(originalDict.Keys)
                                            .Select(id => editedDict[id])
                                            .ToList();

        var deletedComments = originalDict.Keys.Except(editedDict.Keys)
                                               .Select(id => originalDict[id])
                                               .ToList();

        var modifiedComments = originalDict.Keys.Intersect(editedDict.Keys)
                                                .Where(id => !string.Equals(
                                                    originalDict[id].Text,
                                                    editedDict[id].Text,
                                                    StringComparison.Ordinal))
                                                .Select(id => new
                                                {
                                                    Id = id,
                                                    Original = originalDict[id],
                                                    Edited   = editedDict[id]
                                                })
                                                .ToList();

        // -----------------------------------------------------------------
        // 5. Output the results.
        // -----------------------------------------------------------------
        Console.WriteLine("=== Added Comments ===");
        foreach (var c in addedComments)
        {
            Console.WriteLine($"Id: {c.Id}, Author: {c.Author}, Text: {c.Text}");
        }

        Console.WriteLine("\n=== Deleted Comments ===");
        foreach (var c in deletedComments)
        {
            Console.WriteLine($"Id: {c.Id}, Author: {c.Author}, Text: {c.Text}");
        }

        Console.WriteLine("\n=== Modified Comments ===");
        foreach (var m in modifiedComments)
        {
            Console.WriteLine($"Id: {m.Id}");
            Console.WriteLine($"  Original Author: {m.Original.Author}, Text: {m.Original.Text}");
            Console.WriteLine($"  Edited   Author: {m.Edited.Author},   Text: {m.Edited.Text}");
        }

        // -----------------------------------------------------------------
        // 6. (Optional) Demonstrate the Compare API – it creates revisions.
        // -----------------------------------------------------------------
        // The comparison will generate revisions for the comment changes.
        // We do not rely on these revisions for the diff above, but they are
        // useful if you need to inspect the revision collection.
        CompareOptions compareOptions = new CompareOptions
        {
            // Ensure comment changes are not ignored.
            IgnoreComments = false
        };
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

        // Example: list comment‑related revisions.
        Console.WriteLine("\n=== Comment Revisions Detected by Compare ===");
        foreach (Revision rev in original.Revisions)
        {
            // The revision's ParentNode points to the node that changed.
            Node? node = rev.ParentNode;
            if (node != null && node.NodeType == NodeType.Comment)
            {
                Comment revComment = (Comment)node;
                Console.WriteLine($"Revision Type: {rev.RevisionType}, Comment Id: {revComment.Id}, Author: {revComment.Author}");
            }
        }

        // No interactive prompts – the program ends here.
    }

    // Helper method to extract comment data from a document.
    private static List<CommentInfo> GetCommentsInfo(Document doc)
    {
        var list = new List<CommentInfo>();

        // Enumerate all comment nodes in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>())
        {
            // Get the plain text of the comment (trimmed to remove trailing newlines).
            string text = comment.GetText().Trim();

            list.Add(new CommentInfo
            {
                Id = comment.Id,
                Author = comment.Author ?? string.Empty,
                Text = text
            });
        }

        return list;
    }
}
