using System;
using Aspose.Words;
using Aspose.Words.Loading;

class CompareTxtDocuments
{
    static void Main()
    {
        // Paths to the original and edited TXT files.
        string originalPath = "Original.txt";
        string editedPath = "Edited.txt";

        // Load the TXT files into Document objects using TxtLoadOptions.
        // This treats the plain‑text files as Word documents, preserving line breaks.
        Document docOriginal = new Document(originalPath, new TxtLoadOptions());
        Document docEdited   = new Document(editedPath,   new TxtLoadOptions());

        // Ensure both documents have no existing revisions before comparison.
        // (If they had revisions, Compare would throw an exception.)
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the two documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Output the revisions that were created by the comparison.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the compared document as plain text to see the merged result.
        docOriginal.Save("ComparedResult.txt", SaveFormat.Text);
    }
}
