using System;
using Aspose.Words;

// Criteria that matches only deletion revisions.
public class DeletionCriteria : IRevisionCriteria
{
    public bool IsMatch(Revision revision)
    {
        // Keep insertions, format changes, etc.; reject only deletions.
        return revision.RevisionType == RevisionType.Deletion;
    }
}

public class Program
{
    public static void Main()
    {
        // Load the document that contains tracked changes.
        Document doc = new Document("input.docx");

        // Reject all revisions that satisfy the criteria (i.e., all deletions).
        doc.Revisions.Reject(new DeletionCriteria());

        // Save the modified document. Insertions and formatting changes remain intact.
        doc.Save("output.docx");
    }
}
