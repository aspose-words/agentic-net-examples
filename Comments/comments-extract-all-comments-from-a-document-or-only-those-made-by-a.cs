using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class CommentExtractor
{
    /// <summary>
    /// Extracts comments from a DOCX file.
    /// If <paramref name="author"/> is null, all comments are extracted;
    /// otherwise only comments authored by the specified person are extracted.
    /// The extracted comments are written to a new document saved at <paramref name="outputPath"/>.
    /// </summary>
    public static void ExtractComments(string inputPath, string outputPath, string author = null)
    {
        // Load the source document (lifecycle rule: load from file)
        Document sourceDoc = new Document(inputPath);

        // Retrieve all comment nodes in the document (including those in headers/footers)
        NodeCollection commentNodes = sourceDoc.GetChildNodes(NodeType.Comment, true);

        // Create a new blank document to store the extracted comments (lifecycle rule: create)
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Iterate through each comment node
        foreach (Comment comment in commentNodes)
        {
            // If an author filter is supplied, skip comments that don't match
            if (author != null && !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                continue;

            // Write comment details to the result document
            builder.Writeln($"Author: {comment.Author}");
            builder.Writeln($"Date  : {comment.DateTime}");
            builder.Writeln("Text  :");
            builder.Writeln(comment.GetText().Trim());
            builder.Writeln(new string('-', 40));
        }

        // Save the result document (lifecycle rule: save to file)
        resultDoc.Save(outputPath);
    }

    // Example usage
    public static void Main()
    {
        // Path to the source DOCX file
        string sourceFile = @"C:\Docs\Sample.docx";

        // Path where the extracted comments will be saved
        string resultFileAll = @"C:\Docs\AllComments.docx";
        string resultFileAuthor = @"C:\Docs\JohnDoeComments.docx";

        // Extract all comments
        ExtractComments(sourceFile, resultFileAll);

        // Extract only comments made by "John Doe"
        ExtractComments(sourceFile, resultFileAuthor, "John Doe");
    }
}
