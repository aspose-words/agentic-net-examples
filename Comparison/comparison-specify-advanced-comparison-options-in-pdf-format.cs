using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class ComparisonWithAdvancedOptions
{
    static void Main()
    {
        // Paths to the source documents and output folder.
        string MyDir = @"C:\Docs\";
        string ArtifactsDir = @"C:\Output\";

        // Load the original and the edited documents.
        Document docOriginal = new Document(MyDir + "Original.docx");
        Document docEdited   = new Document(MyDir + "Edited.docx");

        // Configure advanced comparison options.
        CompareOptions compareOptions = new CompareOptions();
        // Ignore differences in DrawingML unique IDs (e.g., shapes, charts).
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;
        // Ignore differences in StructuredDocumentTag (SDT) store item IDs.
        compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Set PDF save options, e.g., comply with PDF/A-1b for archival purposes.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the compared document (with revisions) as a PDF file.
        docOriginal.Save(ArtifactsDir + "ComparisonResult.pdf", pdfOptions);
    }
}
