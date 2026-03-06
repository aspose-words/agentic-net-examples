using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docRevised = new Document("Revised.docx");

        // Create CompareOptions and enable advanced options for more precise comparison.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;   // Ignore differences in DrawingML unique IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;   // Ignore differences in StructuredDocumentTag store item IDs.

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now, compareOptions);

        // Set PDF save options – use a PDF/A compliance level that exists in the referenced Aspose.Words version.
        // If PdfA2b is not available, fall back to PdfA2a (PDF/A‑2a) or omit the setting.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Compliance = PdfCompliance.PdfA2b, // <-- unavailable in older versions
            Compliance = PdfCompliance.PdfA2a,   // PDF/A‑2a is supported in all recent versions
            AdditionalTextPositioning = false   // Set true only if you need extra positioning operators.
        };

        // Save the compared document as a PDF file.
        docOriginal.Save("ComparedResult.pdf", pdfOptions);
    }
}
