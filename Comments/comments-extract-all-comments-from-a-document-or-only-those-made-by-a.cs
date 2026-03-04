using System;
using Aspose.Words;

class CommentExtractor
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "Input.docx";

        // Path where the extracted comments will be saved.
        string outputPath = "ExtractedComments.docx";

        // Set to a specific author name to filter, or leave empty/null to get all comments.
        string authorFilter = "John Doe";

        // Load the source document (load rule).
        Document sourceDoc = new Document(inputPath);

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = sourceDoc.GetChildNodes(NodeType.Comment, true);

        // Create a new blank document to hold the extracted comments (create rule).
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Iterate through each comment.
        foreach (Comment comment in commentNodes)
        {
            // If an author filter is specified, skip comments that don't match.
            if (!string.IsNullOrEmpty(authorFilter) &&
                !comment.Author.Equals(authorFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            // Write comment details into the result document.
            builder.Writeln($"Author: {comment.Author}");
            builder.Writeln($"Date  : {comment.DateTime}");
            builder.Writeln("Text  :");
            builder.Writeln(comment.GetText().Trim());
            builder.Writeln(new string('-', 40));
        }

        // Save the document containing the extracted comments (save rule).
        resultDoc.Save(outputPath);
    }
}
