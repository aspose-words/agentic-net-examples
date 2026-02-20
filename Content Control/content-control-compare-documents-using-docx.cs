using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two DOCX documents that contain content controls.
        Document original = new Document("OriginalWithContentControls.docx");
        Document revised = new Document("RevisedWithContentControls.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Do not ignore formatting – we want to see formatting changes inside content controls.
            IgnoreFormatting = false,
            // Compare moves (e.g., moved content controls) as separate revisions.
            CompareMoves = true,
            // Use the revised document as the base for comparison (similar to Word's "Show changes in: New document").
            Target = ComparisonTargetType.New
        };

        // Advanced option: treat content controls with different store item IDs as the same.
        // This is useful when the only difference is the internal ID of a content control.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

        // Perform the comparison. Revisions will be added to the 'original' document.
        original.Compare(revised, "ComparerUser", DateTime.Now, compareOptions);

        // Save the result showing revisions.
        original.Save("ComparisonResult.docx");
    }
}
