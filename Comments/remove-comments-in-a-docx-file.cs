using System;
using Aspose.Words;

class RemoveCommentsExample
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Retrieve all comment nodes in the document (including those in headers/footers).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Remove each comment node from its parent.
        foreach (Comment comment in commentNodes)
        {
            comment.Remove();
        }

        // Save the document without comments.
        doc.Save("Output.docx");
    }
}
