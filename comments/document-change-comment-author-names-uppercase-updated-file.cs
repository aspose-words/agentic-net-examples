using System;
using Aspose.Words;

class CommentAuthorUppercase
{
    static void Main()
    {
        // Load the existing Word document from a file.
        // This uses the Document(string) constructor as required by the lifecycle rules.
        Document doc = new Document("Input.docx");

        // Iterate through all comment nodes in the document.
        // NodeType.Comment is used to retrieve comment nodes.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true))
        {
            // Convert the author name to uppercase and assign it back.
            // The Author property cannot be null, so we guard against it.
            if (!string.IsNullOrEmpty(comment.Author))
                comment.Author = comment.Author.ToUpperInvariant();
        }

        // Save the modified document to a new file.
        // This uses the Save(string) method as defined in the provided rules.
        doc.Save("Output.docx");
    }
}
