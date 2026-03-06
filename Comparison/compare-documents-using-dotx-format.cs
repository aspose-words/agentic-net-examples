using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original document in DOTX format.
        Document docOriginal = new Document("OriginalTemplate.dotx");

        // Load the edited document in DOTX format.
        Document docEdited = new Document("EditedTemplate.dotx");

        // Ensure both documents have no revisions before performing the comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            // "JD" is the author initials that will appear on the revision marks.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Save the result. The saved file contains the tracked changes (revisions).
        docOriginal.Save("ComparisonResult.docx");
    }
}
