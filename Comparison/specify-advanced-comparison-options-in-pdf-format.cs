using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class AdvancedComparisonToPdf
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Configure advanced comparison options.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;   // Ignore DrawingML unique IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;   // Ignore StructuredDocumentTag store item IDs.

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

        // Set PDF save options (example: PDF/A-1b compliance).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the comparison result as a PDF file.
        original.Save("ComparisonResult.pdf", pdfOptions);
    }
}
