using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file that contains comments.
        string inputPath = @"C:\Docs\SourceWithComments.docx";

        // Path where the cleaned DOCX file will be saved.
        string outputPath = @"C:\Docs\SourceWithoutComments.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve all comment nodes in the document (including those in headers/footers).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Remove each comment node from its parent. Iterate backwards to avoid collection modification issues.
        for (int i = commentNodes.Count - 1; i >= 0; i--)
        {
            commentNodes[i].Remove();
        }

        // Save the modified document back to disk.
        doc.Save(outputPath);
    }
}
