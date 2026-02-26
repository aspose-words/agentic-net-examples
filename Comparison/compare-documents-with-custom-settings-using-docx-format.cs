using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the source documents and the output folder.
        string dataDir = @"C:\Docs\";
        string outputDir = @"C:\Output\";

        // Load the original and the edited documents.
        Document docOriginal = new Document(dataDir + "Original.docx");
        Document docEdited   = new Document(dataDir + "Edited.docx");

        // Configure custom comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the character level.
            Granularity = Granularity.CharLevel,

            // Use the edited document as the base for comparison (equivalent to Word's "Show changes in: New document").
            Target = ComparisonTargetType.New,

            // Example of ignoring specific element types.
            IgnoreFormatting = true,
            IgnoreComments   = true,
            IgnoreTables     = false,
            IgnoreFootnotes  = false,
            IgnoreTextboxes  = false,
            IgnoreHeadersAndFooters = false,

            // Advanced options – ignore DrawingML unique IDs to avoid false positives.
            // The AdvancedOptions property is read‑only, but its members are settable.
            // This follows the provided API contract.
        };
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the result (the original document now contains revision marks).
        docOriginal.Save(outputDir + "ComparisonResult.docx");
    }
}
