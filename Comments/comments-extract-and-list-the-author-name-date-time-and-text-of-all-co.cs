using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ExtractComments
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment and output its details.
        foreach (Comment comment in commentNodes)
        {
            // Author of the comment.
            string author = comment.Author;

            // Date and time when the comment was made.
            DateTime dateTime = comment.DateTime;

            // Text content of the comment (trimmed to remove extra whitespace).
            string text = comment.GetText().Trim();

            Console.WriteLine($"Author: {author}");
            Console.WriteLine($"DateTime: {dateTime}");
            Console.WriteLine($"Text: {text}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
