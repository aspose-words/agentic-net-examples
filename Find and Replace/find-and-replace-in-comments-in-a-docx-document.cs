using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class CommentFindReplace
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        const string findText = "_FullName_";
        const string replaceText = "John Doe";

        // Iterate through all comment nodes in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true))
        {
            // Perform a simple find-and-replace within the comment's range.
            // The Replace method is case‑insensitive by default.
            comment.Range.Replace(findText, replaceText);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
