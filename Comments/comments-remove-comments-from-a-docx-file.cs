using System;
using Aspose.Words;

public class CommentRemover
{
    /// <summary>
    /// Removes all comments from a DOCX file.
    /// </summary>
    /// <param name="inputFile">Path to the source DOCX file.</param>
    /// <param name="outputFile">Path where the cleaned DOCX will be saved.</param>
    public static void RemoveComments(string inputFile, string outputFile)
    {
        // Load the document.
        Document doc = new Document(inputFile);

        // Retrieve all comment nodes in the document.
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

        // Remove each comment node from its parent. Iterate backwards to avoid index shifting.
        for (int i = comments.Count - 1; i >= 0; i--)
        {
            comments[i].Remove();
        }

        // Save the modified document.
        doc.Save(outputFile);
    }

    // Example usage.
    public static void Main()
    {
        string sourcePath = @"C:\Docs\input.docx";
        string destinationPath = @"C:\Docs\output_without_comments.docx";

        RemoveComments(sourcePath, destinationPath);
        Console.WriteLine("Comments removed and document saved to: " + destinationPath);
    }
}
