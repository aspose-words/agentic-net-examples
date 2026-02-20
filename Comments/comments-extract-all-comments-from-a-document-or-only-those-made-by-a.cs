using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

class ExtractComments
{
    static void Main(string[] args)
    {
        // Path to the input DOCX file.
        string inputPath = "InputDocument.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve all comment nodes in the document.
        // The Document.GetChildNodes method returns a NodeCollection; we filter it to Comment objects.
        List<Comment> allComments = doc.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .ToList();

        // Example: Extract all comments.
        Console.WriteLine("All comments in the document:");
        foreach (Comment comment in allComments)
        {
            Console.WriteLine($"Author: {comment.Author}");
            Console.WriteLine($"Date: {comment.DateTime}");
            Console.WriteLine($"Text: {comment.GetText().Trim()}");
            Console.WriteLine(new string('-', 40));
        }

        // Example: Extract only comments made by a specific author.
        string targetAuthor = "John Doe"; // Change to the desired author name.
        IEnumerable<Comment> authorComments = allComments.Where(c => c.Author == targetAuthor);

        Console.WriteLine($"\nComments by author \"{targetAuthor}\":");
        foreach (Comment comment in authorComments)
        {
            Console.WriteLine($"Date: {comment.DateTime}");
            Console.WriteLine($"Text: {comment.GetText().Trim()}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
