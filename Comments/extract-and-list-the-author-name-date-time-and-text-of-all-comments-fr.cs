using System;
using Aspose.Words;

class ExtractComments
{
    static void Main()
    {
        // Path to the DOCX file.
        string docPath = @"C:\Path\To\Your\Document.docx";

        // Load the document.
        Document doc = new Document(docPath);

        // Retrieve all comment nodes (including replies).
        var commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        foreach (Comment comment in commentNodes)
        {
            string author = comment.Author;
            DateTime dateTime = comment.DateTime;
            string text = comment.GetText().Trim();

            Console.WriteLine($"Author: {author}");
            Console.WriteLine($"Date & Time: {dateTime}");
            Console.WriteLine($"Comment Text: {text}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
