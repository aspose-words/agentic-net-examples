using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment and output its details.
        foreach (Comment comment in commentNodes)
        {
            Console.WriteLine($"Author: {comment.Author}");
            Console.WriteLine($"DateTime: {comment.DateTime}");
            Console.WriteLine($"Text: {comment.GetText().Trim()}");
            Console.WriteLine();
        }
    }
}
