using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";
        string resultPath   = "Compared.docx";

        // Load the documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The CompareOptions are left at defaults,
            // which means comment changes are NOT ignored.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the document that now contains the revisions (optional).
        docOriginal.Save(resultPath);

        // Iterate through all revisions and output information about comment changes.
        foreach (Revision revision in docOriginal.Revisions)
        {
            // Filter only comment revisions.
            if (revision.ParentNode.NodeType == NodeType.Comment)
            {
                // Determine the type of change (Insertion = added, Deletion = removed).
                string changeType = revision.RevisionType == RevisionType.Insertion ? "Added"
                                 : revision.RevisionType == RevisionType.Deletion ? "Deleted"
                                 : "Modified";

                // Get the comment text. For deletions the node still exists in the revision.
                string commentText = revision.ParentNode.GetText().Trim();

                Console.WriteLine($"Comment {changeType}: Author = {revision.Author}, Text = \"{commentText}\"");
            }
        }
    }
}
