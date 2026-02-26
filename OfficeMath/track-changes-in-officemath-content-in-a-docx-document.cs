using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class TrackOfficeMathChanges
{
    static void Main()
    {
        // Load the original document that contains OfficeMath equations.
        Document originalDoc = new Document("Original.docx");

        // Load the edited version of the document where the OfficeMath content has been changed.
        Document editedDoc = new Document("Edited.docx");

        // Ensure that neither document already contains revisions; otherwise Compare will throw.
        if (originalDoc.HasRevisions || editedDoc.HasRevisions)
            throw new InvalidOperationException("Documents must not have existing revisions before comparison.");

        // Compare the two documents. This will create insertion/deletion revisions for any changes,
        // including additions, removals, or replacements of OfficeMath nodes.
        originalDoc.Compare(editedDoc, "John Doe", DateTime.Now);

        // Iterate over the generated revisions and output basic information.
        foreach (Revision rev in originalDoc.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
        }

        // Save the original document now containing tracked changes.
        originalDoc.Save("TrackedOfficeMathChanges.docx");
    }
}
