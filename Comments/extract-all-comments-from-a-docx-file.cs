using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file (uses the Document constructor – the provided load rule)
        Document doc = new Document("input.docx");

        // Retrieve all comment nodes in the document (including those in headers/footers)
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment and extract its details
        foreach (Comment comment in commentNodes)
        {
            // The comment's text includes all its child paragraphs; Trim removes trailing line breaks
            string commentText = comment.GetText().Trim();

            // Output comment information
            Console.WriteLine($"Comment Id: {comment.Id}");
            Console.WriteLine($"Author: {comment.Author}");
            Console.WriteLine($"Date: {comment.DateTime}");
            Console.WriteLine($"Text: {commentText}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
