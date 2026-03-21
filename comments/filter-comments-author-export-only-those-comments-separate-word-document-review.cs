using System;
using System.Linq;
using Aspose.Words;

class FilterCommentsByAuthor
{
    static void Main()
    {
        // Create a source document with sample comments.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is a sample document.");

        // Add a comment by the target author.
        Comment comment1 = new Comment(sourceDoc, "John Doe", "JD", DateTime.Now);
        comment1.Paragraphs.Add(new Paragraph(sourceDoc));
        comment1.FirstParagraph.Runs.Add(new Run(sourceDoc, "Comment from John Doe."));
        srcBuilder.CurrentParagraph.AppendChild(comment1);

        // Add a comment by another author.
        Comment comment2 = new Comment(sourceDoc, "Jane Smith", "JS", DateTime.Now);
        comment2.Paragraphs.Add(new Paragraph(sourceDoc));
        comment2.FirstParagraph.Runs.Add(new Run(sourceDoc, "Comment from Jane Smith."));
        srcBuilder.CurrentParagraph.AppendChild(comment2);

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
