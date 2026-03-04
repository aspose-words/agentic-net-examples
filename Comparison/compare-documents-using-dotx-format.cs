using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDotxTemplates
{
    static void Main()
    {
        // Load the original DOTX template.
        Document docOriginal = new Document("OriginalTemplate.dotx");

        // Load the edited DOTX template to compare against.
        Document docEdited = new Document("EditedTemplate.dotx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the two documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the comparison result (with revisions) as a DOTX file.
        docOriginal.Save("ComparisonResult.dotx");
    }
}
