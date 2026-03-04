using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocm
{
    static void Main()
    {
        // Load the original DOCM document.
        Document docOriginal = new Document("Original.docm");

        // Load the edited DOCM document to compare against.
        Document docEdited = new Document("Edited.docm");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options (optional). Here we keep formatting changes.
            CompareOptions compareOptions = new CompareOptions
            {
                // Set to false to track formatting changes; change to true to ignore them.
                IgnoreFormatting = false,
                // Use the edited document as the target for comparison.
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);
        }

        // Save the resulting document (which now contains revision marks) as DOCM.
        docOriginal.Save("Compared.docm");
    }
}
