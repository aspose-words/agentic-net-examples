using System;
using System.Collections.Generic;
using Aspose.Words;

class ExtractComments
{
    static void Main()
    {
        // Load the DOCX file (uses the provided Document(string) constructor)
        Document doc = new Document("Input.docx");

        // Retrieve all comment nodes in the document (including nested replies)
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Prepare a list to hold extracted comment information
        List<string> extractedComments = new List<string>();

        // Iterate through each comment node
        foreach (Comment comment in commentNodes)
        {
            // Build a readable representation: Author, Date, and comment text
            string commentInfo = $"Author: {comment.Author}, Date: {comment.DateTime}, Text: {comment.GetText().Trim()}";
            extractedComments.Add(commentInfo);
        }

        // Output the extracted comments to the console
        Console.WriteLine("Extracted Comments:");
        foreach (string info in extractedComments)
        {
            Console.WriteLine(info);
        }

        // (Optional) Save the list of comments to a text file for further processing
        // Uses the provided Document.Save(string) rule if you need to create a document with the comments.
        // Here we simply write to a plain text file.
        System.IO.File.WriteAllLines("Comments.txt", extractedComments);
    }
}
