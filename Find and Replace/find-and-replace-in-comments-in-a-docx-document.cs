using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class CommentFindReplace
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Optional: configure find/replace options (e.g., case‑insensitive, whole‑word).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate over all comment nodes in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true))
        {
            // Perform the replace operation within the comment's range.
            comment.Range.Replace(findText, replaceText, options);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
