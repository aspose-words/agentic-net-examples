using System;
using System.Collections.Generic;
using Aspose.Words;

class RejectFormattingRevisions
{
    static void Main()
    {
        // Load the document that contains tracked changes.
        Document doc = new Document("Input.docx");

        // Collect all revisions that are formatting changes.
        List<Revision> formatRevisions = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                formatRevisions.Add(rev);
        }

        // Reject only the formatting revisions; insertions and deletions remain untouched.
        foreach (Revision rev in formatRevisions)
        {
            rev.Reject(); // Uses Revision.Reject method.
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
