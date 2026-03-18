using System;
using System.Linq;
using Aspose.Words;

class FilterCommentsByAuthor
{
    static void Main()
    {
        // Load the source document that contains comments.
        Document sourceDoc = new Document("InputDocument.docx");

        // Create a new blank document that will hold the filtered comments.
        Document filteredDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(filteredDoc);

        // Title for the exported comments.
        builder.Writeln("Filtered Comments");
        builder.Writeln("-----------------");
        builder.Writeln();

        // Retrieve all comment nodes from the source document.
        NodeCollection allComments = sourceDoc.GetChildNodes(NodeType.Comment, true);

        // Define the author whose comments we want to export.
        const string targetAuthor = "John Doe";

        // Iterate through each comment, copy only those authored by the target author.
        foreach (Comment comment in allComments.OfType<Comment>())
        {
            if (string.Equals(comment.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            {
                // Write comment metadata.
                builder.Writeln($"Author : {comment.Author}");
                builder.Writeln($"Date   : {comment.DateTime}");
                builder.Writeln($"Text   : {comment.GetText().Trim()}");
                builder.Writeln(); // Add a blank line between comments.
            }
        }

        // Save the new document containing only the filtered comments.
        filteredDoc.Save("FilteredComments.docx");
    }
}
