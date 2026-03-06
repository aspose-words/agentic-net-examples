using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDotm
{
    static void Main()
    {
        // Load the original macro-enabled template (.dotm).
        Document docOriginal = new Document("OriginalTemplate.dotm");

        // Load the edited version of the template.
        Document docEdited = new Document("EditedTemplate.dotm");

        // Ensure both documents have no existing revisions before comparing.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options as needed.
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: track formatting changes; set to true to ignore them.
                IgnoreFormatting = false,
                // Use the edited document as the target for comparison.
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);
        }

        // Save the comparison result as a macro-enabled template.
        docOriginal.Save("ComparisonResult.dotm");
    }
}
