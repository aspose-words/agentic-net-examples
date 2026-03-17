using System;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Replacing;

class CommentToFootnote
{
    static void Main()
    {
        // Load an existing document that contains comments.
        // Replace the path with the actual location of your source file.
        Document doc = new Document(@"C:\Docs\InputWithComments.docx");

        // Create a DocumentBuilder for inserting footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Retrieve all comment nodes in the document (including replies).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate over each comment.
        foreach (Comment comment in commentNodes)
        {
            // Get the full text of the comment (including paragraph breaks).
            string commentText = comment.GetText();

            // Position the builder right after the comment node.
            // This will place the footnote marker immediately after the comment anchor.
            builder.MoveTo(comment);
            builder.InsertFootnote(FootnoteType.Footnote, commentText);
        }

        // Save the modified document. Adjust the output path as needed.
        doc.Save(@"C:\Docs\OutputWithFootnotes.docx");
    }
}
