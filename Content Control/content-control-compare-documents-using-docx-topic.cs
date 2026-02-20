using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace CompareContentControls
{
    class Program
    {
        static void Main()
        {
            // Load the original document that contains content controls.
            Document originalDoc = new Document("Original.docx");

            // Load the edited document that also contains content controls.
            Document editedDoc = new Document("Edited.docx");

            // Configure comparison options.
            CompareOptions compareOptions = new CompareOptions
            {
                // Track changes at the word level (default is character level).
                Granularity = Granularity.WordLevel,

                // Do not ignore formatting differences.
                IgnoreFormatting = false,

                // Do not ignore comments.
                IgnoreComments = false,

                // Use the edited document as the base for comparison.
                Target = ComparisonTargetType.New
            };

            // Advanced options: consider differences in StructuredDocumentTag (content control) store item IDs.
            compareOptions.AdvancedOptions.IgnoreStoreItemId = false;

            // Perform the comparison. Revisions will be added to the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now, compareOptions);

            // Save the result with revisions visible.
            originalDoc.Save("ComparedResult.docx");
        }
    }
}
