using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Replacing;

class OfficeMathComparison
{
    static void Main()
    {
        // Load the original and the revised documents.
        Document originalDoc = new Document("Original.docx");
        Document revisedDoc = new Document("Revised.docx");

        // Set up comparison options (default options are sufficient for OfficeMath comparison).
        CompareOptions compareOptions = new CompareOptions();

        // Perform the comparison. The original document will receive revision marks.
        originalDoc.Compare(revisedDoc, "Comparer", DateTime.Now, compareOptions);

        // Count revisions that affect OfficeMath nodes.
        int officeMathRevisions = 0;
        foreach (Revision rev in originalDoc.Revisions)
        {
            if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.OfficeMath)
                officeMathRevisions++;
        }

        // Output the result.
        Console.WriteLine($"Number of OfficeMath revisions: {officeMathRevisions}");

        // Save the document with revisions for visual inspection.
        originalDoc.Save("ComparisonResult.docx");
    }
}
