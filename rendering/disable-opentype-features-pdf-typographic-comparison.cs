using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

namespace OpenTypeFeatureDisableExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Docs\ResultDocument.pdf";

            // Load the existing Word document.
            Document doc = new Document(inputPath);

            // Disable OpenType font formatting features for the whole document.
            // This ensures that advanced OpenType features (ligatures, contextual forms, etc.) are ignored.
            doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

            // Create PDF save options. No special options are required for this scenario,
            // but the object is instantiated to follow the standard save pattern.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the document as PDF using the configured options.
            doc.Save(outputPath, pdfOptions);
        }
    }
}
