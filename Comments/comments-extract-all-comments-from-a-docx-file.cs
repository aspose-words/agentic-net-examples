using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Store the extracted comment texts.
        List<string> commentTexts = new List<string>();

        // Retrieve all comment nodes (including replies) from the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment comment in commentNodes)
        {
            // Get the plain text of the comment and trim any trailing line breaks.
            string text = comment.GetText().Trim();
            commentTexts.Add(text);
        }

        // Output each comment to the console.
        foreach (string text in commentTexts)
        {
            Console.WriteLine(text);
        }
    }
}
