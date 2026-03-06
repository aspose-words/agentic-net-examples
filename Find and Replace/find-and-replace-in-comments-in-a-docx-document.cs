using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class CommentFindReplace
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options (case‑insensitive in this example).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate through all comment nodes in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true))
        {
            // Replace the target text inside the comment's range.
            comment.Range.Replace("_FullName_", "John Doe", options);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
