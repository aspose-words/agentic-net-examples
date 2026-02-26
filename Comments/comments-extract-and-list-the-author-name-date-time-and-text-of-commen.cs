using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

class ExtractKsComments
{
    static void Main()
    {
        // Load the DOCX document.
        string inputPath = "input.docx";
        Document doc = new Document(inputPath); // uses the Document(string) constructor

        // Collect comments authored by "ks".
        List<string> ksComments = new List<string>();

        // Get all comment nodes in the document (including those in headers/footers).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        foreach (Comment comment in commentNodes)
        {
            if (string.Equals(comment.Author, "ks", StringComparison.OrdinalIgnoreCase))
            {
                // Build a readable line with author, date/time and comment text.
                string line = $"Author: {comment.Author}, DateTime: {comment.DateTime}, Text: {comment.GetText().Trim()}";
                ksComments.Add(line);
            }
        }

        // Output the collected comments.
        Console.WriteLine("Comments authored by 'ks':");
        foreach (string line in ksComments)
        {
            Console.WriteLine(line);
        }
    }
}
