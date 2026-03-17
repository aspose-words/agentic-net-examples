using System;
using System.Linq;
using Aspose.Words;

class DeleteCommentsByAuthor
{
    static void Main()
    {
        // Path to the source document.
        const string inputPath = @"C:\Docs\Input.docx";
        // Path where the modified document will be saved.
        const string outputPath = @"C:\Docs\Output.docx";

        // Load the document (lifecycle rule: load).
        Document doc = new Document(inputPath);

        // Author whose comments should be removed.
        const string authorToDelete = "John Doe";

        // Get all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate backwards so that removal does not affect the loop index.
        for (int i = commentNodes.Count - 1; i >= 0; i--)
        {
            Comment comment = (Comment)commentNodes[i];

            // Check the author of the comment.
            if (string.Equals(comment.Author, authorToDelete, StringComparison.OrdinalIgnoreCase))
            {
                // Remove the associated comment range start node, if it exists.
                CommentRangeStart rangeStart = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                    .Cast<CommentRangeStart>()
                    .FirstOrDefault(r => r.Id == comment.Id);
                rangeStart?.Remove();

                // Remove the associated comment range end node, if it exists.
                CommentRangeEnd rangeEnd = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                    .Cast<CommentRangeEnd>()
                    .FirstOrDefault(r => r.Id == comment.Id);
                rangeEnd?.Remove();

                // Finally, remove the comment itself.
                comment.Remove();
            }
        }

        // Save the modified document (lifecycle rule: save).
        doc.Save(outputPath);
    }
}
