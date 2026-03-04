using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("Comments.docx");

        // Get all comment nodes in the document (including those inside other nodes).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate over top‑level comments only (skip replies).
        foreach (Comment comment in commentNodes.OfType<Comment>().Where(c => c.Ancestor == null))
        {
            // Extract required information.
            string author = comment.Author;
            DateTime dateTime = comment.DateTime;
            string text = comment.GetText().Trim();

            // Output the comment details.
            Console.WriteLine($"Author: {author}");
            Console.WriteLine($"Date & Time: {dateTime}");
            Console.WriteLine($"Text: {text}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
