using System;
using Aspose.Words;

class RemoveCommentsFromDocx
{
    static void Main()
    {
        // Input and output file paths.
        string inputPath = @"C:\Docs\InputDocument.docx";
        string outputPath = @"C:\Docs\OutputDocument.docx";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Remove all comment nodes.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment comment in commentNodes)
        {
            comment.Remove();
        }

        // Remove comment range start markers.
        NodeCollection rangeStartNodes = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        foreach (CommentRangeStart start in rangeStartNodes)
        {
            start.Remove();
        }

        // Remove comment range end markers.
        NodeCollection rangeEndNodes = doc.GetChildNodes(NodeType.CommentRangeEnd, true);
        foreach (CommentRangeEnd end in rangeEndNodes)
        {
            end.Remove();
        }

        // Save the cleaned document.
        doc.Save(outputPath);
    }
}
