using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class AdvancedComparisonDemo
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docModified = new Document("Modified.docx");

        // Create a CompareOptions instance and configure advanced options.
        CompareOptions compareOptions = new CompareOptions();

        // Ignore differences in DrawingML unique IDs (e.g., shape IDs).
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;

        // Ignore differences in StructuredDocumentTag (SDT) store item IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

        // Perform the comparison. The revisions will be added to docOriginal.
        docOriginal.Compare(docModified, "Comparer", DateTime.Now, compareOptions);

        // Save the result document with revisions to a DOCX file.
        docOriginal.Save("ComparisonResult.docx");
    }
}
