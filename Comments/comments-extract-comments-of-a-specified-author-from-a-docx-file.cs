using System;
using Aspose.Words;

class ExtractAuthorComments
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the extracted comments will be saved.
        string outputPath = "author_comments.docx";

        // The author whose comments we want to extract.
        string targetAuthor = "John Doe";

        // Load the existing document.
        Document sourceDoc = new Document(inputPath);

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = sourceDoc.GetChildNodes(NodeType.Comment, true);

        // Create a new blank document to store the extracted comments.
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Header for the extracted comments.
        builder.Writeln($"Comments authored by \"{targetAuthor}\":");
        builder.Writeln();

        int extractedCount = 0;

        // Iterate through all comments and copy those matching the target author.
        foreach (Comment comment in commentNodes)
        {
            if (comment.Author == targetAuthor)
            {
                extractedCount++;
                // Write the comment text to the new document.
                builder.Writeln($"- {comment.GetText().Trim()}");
            }
        }

        // If no comments were found, indicate that in the output document.
        if (extractedCount == 0)
        {
            builder.Writeln("No comments found for the specified author.");
        }

        // Save the document containing the extracted comments.
        resultDoc.Save(outputPath);
    }
}
