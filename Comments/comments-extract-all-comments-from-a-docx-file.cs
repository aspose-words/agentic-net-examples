using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file.
        string inputPath = "input.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Extract all comments.
        List<string> commentTexts = ExtractComments(doc);

        // Output the comments.
        foreach (string text in commentTexts)
            Console.WriteLine(text);
    }

    // Returns a list containing the text of each comment in the document.
    static List<string> ExtractComments(Document doc)
    {
        var comments = new List<string>();

        // Retrieve all comment nodes from the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment and collect its text.
        foreach (Comment comment in commentNodes)
        {
            // The comment's text is stored in its range.
            string text = comment.Range.Text.Trim();
            comments.Add(text);
        }

        return comments;
    }
}
