using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareIgnoreCase
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options to ignore case changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreCaseChanges = true   // Makes the comparison case‑insensitive.
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the result document containing the revisions.
        docOriginal.Save("ComparisonResult.docx");
    }
}
