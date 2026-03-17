using System;
using Aspose.Words; // Comment class resides directly in Aspose.Words namespace
using System.Linq;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Define the date threshold. Comments older than this date will be removed.
        DateTime threshold = new DateTime(2023, 1, 1);

        // Retrieve all comment nodes (including replies) from the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate backwards so that removing a comment does not affect the loop indices.
        for (int i = commentNodes.Count - 1; i >= 0; i--)
        {
            Comment comment = (Comment)commentNodes[i];

            // If the comment's DateTime is earlier than the threshold, remove it.
            if (comment.DateTime < threshold)
            {
                // Removing the comment also removes its associated range start/end nodes.
                comment.Remove();
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
