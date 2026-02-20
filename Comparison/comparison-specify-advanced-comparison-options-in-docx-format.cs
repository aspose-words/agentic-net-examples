using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class CompareAdvanced
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document docOriginal = new Document("Original.docx");
        Document docRevised = new Document("Revised.docx");

        // Set up comparison options with advanced settings.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level (instead of character level).
            Granularity = Granularity.WordLevel,
            // Ignore formatting differences.
            IgnoreFormatting = true,
            // Treat case changes as insignificant.
            IgnoreCaseChanges = true,
            // Include move tracking.
            CompareMoves = true,
            // Use the revised document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // Advanced options that have no equivalent in Microsoft Word.
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;   // Ignore DrawingML unique IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;   // Ignore StructuredDocumentTag store item IDs.

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now, compareOptions);

        // Save the result as a DOCX file with strict OOXML compliance.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Strict,
            // Ensure fields are updated before saving (optional, but keeps the document consistent).
            UpdateFields = true
        };

        docOriginal.Save("ComparisonResult.docx", saveOptions);
    }
}
