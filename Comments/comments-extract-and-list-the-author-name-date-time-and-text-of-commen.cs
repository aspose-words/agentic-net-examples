using System;
using Aspose.Words;

class ExtractKsComments
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment and process those authored by "ks".
        foreach (Comment comment in commentNodes)
        {
            if (string.Equals(comment.Author, "ks", StringComparison.OrdinalIgnoreCase))
            {
                // Output author name.
                Console.WriteLine($"Author: {comment.Author}");

                // Output the date and time the comment was made.
                Console.WriteLine($"DateTime: {comment.DateTime}");

                // Output the comment text. Trim to remove leading/trailing whitespace.
                Console.WriteLine($"Text: {comment.GetText().Trim()}");

                Console.WriteLine(new string('-', 40));
            }
        }
    }
}
