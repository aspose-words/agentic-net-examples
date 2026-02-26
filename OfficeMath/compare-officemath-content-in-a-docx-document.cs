using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class OfficeMathComparison
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Perform the comparison. The revisions will be added to docOriginal.
        CompareOptions compareOptions = new CompareOptions();
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Count how many revisions affect OfficeMath nodes.
        int officeMathRevisions = 0;
        foreach (Revision rev in docOriginal.Revisions)
        {
            if (rev.ParentNode.NodeType == NodeType.OfficeMath)
                officeMathRevisions++;
        }

        Console.WriteLine($"Total revisions: {docOriginal.Revisions.Count}");
        Console.WriteLine($"OfficeMath revisions: {officeMathRevisions}");

        // Save the document that now contains the revision marks.
        docOriginal.Save("Compared.docx");
    }
}
