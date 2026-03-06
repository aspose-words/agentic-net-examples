using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDotmDocuments
{
    static void Main()
    {
        // Path to the folder that contains the DOTM files.
        string docsPath = @"C:\Docs\";

        // Load the original and the edited documents (both are DOTM files).
        Document originalDoc = new Document(docsPath + "Original.dotm");
        Document editedDoc   = new Document(docsPath + "Edited.dotm");

        // The Compare method requires that both documents have no existing revisions.
        if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
        {
            // Perform a simple comparison. The author name and the current time are stored
            // in each revision that is created.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
        }

        // Save the result of the comparison back to a DOTM file.
        originalDoc.Save(docsPath + "ComparisonResult.dotm");
    }
}
