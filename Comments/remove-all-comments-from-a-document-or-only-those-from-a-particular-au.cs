using System;
using Aspose.Words;

class RemoveCommentsExample
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\input.docx";

        // Path where the result will be saved.
        string outputFile = @"C:\Docs\output.docx";

        // Set to a specific author to delete only that author's comments,
        // or leave null/empty to delete all comments.
        string authorToDelete = null; // e.g. "John Doe";

        // Load the document (lifecycle rule: load).
        Document doc = new Document(inputFile);

        // Retrieve all comment nodes in the document (including those in headers/footers).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate backwards so that removal does not affect the collection indexing.
        for (int i = commentNodes.Count - 1; i >= 0; i--)
        {
            Comment comment = (Comment)commentNodes[i];

            // If no author filter is set, or the comment's author matches the filter, remove it.
            if (string.IsNullOrEmpty(authorToDelete) || comment.Author == authorToDelete)
                comment.Remove();
        }

        // Save the modified document (lifecycle rule: save).
        doc.Save(outputFile);
    }
}
