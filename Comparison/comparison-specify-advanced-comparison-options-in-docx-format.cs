using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document docOriginal = new Document("Original.docx");
        Document docRevised = new Document("Revised.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions();

        // Advanced options: ignore differences that are not meaningful for most scenarios.
        // - Ignore DrawingML unique IDs (e.g., shape IDs that change on each save).
        // - Ignore StructuredDocumentTag (SDT) store item IDs.
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

        // Set which document is treated as the base during comparison.
        // This mimics Word's "Show changes in" option.
        compareOptions.Target = ComparisonTargetType.New;

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document (with revisions) in DOCX format.
        docOriginal.Save("ComparisonResult.docx");
    }
}
