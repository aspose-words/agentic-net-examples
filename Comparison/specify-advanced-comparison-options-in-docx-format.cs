using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class AdvancedComparisonDemo
{
    static void Main()
    {
        // Load the original and the revised documents.
        Document original = new Document(@"C:\Docs\Original.docx");
        Document revised = new Document(@"C:\Docs\Revised.docx");

        // Create a CompareOptions instance and configure advanced options.
        CompareOptions compareOptions = new CompareOptions();

        // Ignore differences in DrawingML unique IDs (e.g., shape IDs).
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;

        // Ignore differences in StructuredDocumentTag (SDT) store item IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

        // Perform the comparison. The revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result to a new DOCX file.
        original.Save(@"C:\Docs\ComparisonResult.docx");
    }
}
