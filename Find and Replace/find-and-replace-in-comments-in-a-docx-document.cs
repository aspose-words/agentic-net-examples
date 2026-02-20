using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceInComments
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        string oldText = "PLACEHOLDER";
        string newText = "ActualValue";

        // Iterate through all comments in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true))
        {
            // Perform a find/replace operation within the comment's range.
            comment.Range.Replace(oldText, newText, new FindReplaceOptions());
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
