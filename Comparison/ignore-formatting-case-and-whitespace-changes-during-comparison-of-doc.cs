using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options to ignore formatting, case changes and whitespace differences.
        // IgnoreFormatting: skips formatting changes.
        // IgnoreCaseChanges: makes the comparison case‑insensitive.
        // Granularity set to CharLevel to detect changes at the character level (including whitespace).
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting   = true,
            IgnoreCaseChanges  = true,
            Granularity        = Granularity.CharLevel,
            // Other flags can remain false (default) to keep their differences tracked.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Optionally accept all revisions so the original document becomes identical to the edited one.
        docOriginal.Revisions.AcceptAll();

        // Save the result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
