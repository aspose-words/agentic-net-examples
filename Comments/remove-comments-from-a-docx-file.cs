using System;
using Aspose.Words;
using Aspose.Words.Tables;

class RemoveCommentsFromDocx
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

        // Save the modified document without comments.
        doc.Save("Output.docx");
    }
}
