using System;
using Aspose.Words;

class RemoveCommentsExample
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Retrieve all comment nodes in the document (including those in headers/footers).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Remove each comment node. Iterate backwards to avoid index shifting after removal.
        for (int i = commentNodes.Count - 1; i >= 0; i--)
        {
            commentNodes[i].Remove();
        }

        // Save the document without comments.
        doc.Save("Output.docx");
    }
}
