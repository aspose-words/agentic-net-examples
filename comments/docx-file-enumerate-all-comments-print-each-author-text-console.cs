using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file from disk.
        // Replace "input.docx" with the path to your document.
        Document doc = new Document("input.docx");

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment and output its author and text.
        foreach (Comment comment in commentNodes)
        {
            // The Author property holds the name of the comment's author.
            string author = comment.Author;

            // GetText() returns the comment's content (including any child nodes).
            // Trim to remove leading/trailing whitespace and line breaks.
            string text = comment.GetText().Trim();

            Console.WriteLine($"Author: {author}, Text: {text}");
        }
    }
}
