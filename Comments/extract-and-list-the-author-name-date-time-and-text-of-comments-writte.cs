using System;
using System.Linq;
using Aspose.Words;

class ExtractCommentsByAuthor
{
    static void Main()
    {
        // Path to the DOCX file to be processed.
        const string docPath = @"C:\Path\To\Your\Document.docx";

        // Load the document using the provided Document(string) constructor.
        Document doc = new Document(docPath);

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Filter comments authored by "ks" and output their details.
        foreach (Comment comment in commentNodes.OfType<Comment>()
                                                .Where(c => c.Author == "ks"))
        {
            // Trim the comment text to remove leading/trailing whitespace.
            string commentText = comment.GetText().Trim();

            // Output author, date/time, and comment text.
            Console.WriteLine($"Author   : {comment.Author}");
            Console.WriteLine($"DateTime : {comment.DateTime}");
            Console.WriteLine($"Text     : {commentText}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
