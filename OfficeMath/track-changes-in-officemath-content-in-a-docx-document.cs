using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document that contains OfficeMath content.
        // The Document constructor with a file path is the provided load rule.
        Document doc = new Document("OfficeMathInput.docx");

        // Begin tracking revisions. All subsequent node insertions or deletions
        // will be recorded as revision changes with the specified author.
        doc.StartTrackRevisions("Reviewer");

        // Locate the first OfficeMath node in the document.
        // OfficeMath objects can only be children of Paragraph nodes.
        OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (officeMath != null)
        {
            // Deleting the OfficeMath node creates a deletion revision.
            // Aspose.Words tracks node deletions, which we can later accept or reject.
            officeMath.Remove();
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the modified document. The Save method is the provided save rule.
        doc.Save("OfficeMathTrackedChanges.docx");
    }
}
