using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file (replace with your actual file path)
        Document doc = new Document("input.docx");

        // Retrieve all comment nodes in the document
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Filter comments authored by "ks"
        var ksComments = commentNodes
            .OfType<Comment>()
            .Where(c => c.Author.Equals("ks", StringComparison.OrdinalIgnoreCase));

        // List author, date/time, and comment text for each matching comment
        foreach (Comment comment in ksComments)
        {
            string author = comment.Author;
            DateTime dateTime = comment.DateTime;
            string text = comment.GetText().Trim();

            Console.WriteLine($"Author: {author}");
            Console.WriteLine($"DateTime: {dateTime}");
            Console.WriteLine($"Text: {text}");
            Console.WriteLine();
        }
    }
}
