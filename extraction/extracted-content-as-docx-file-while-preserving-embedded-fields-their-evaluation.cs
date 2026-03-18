using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PreserveFieldsExample
{
    static void Main()
    {
        // Path to the source document that contains the content to be extracted.
        string sourcePath = @"C:\Docs\Source.docx";

        // Path where the extracted content will be saved as a DOCX file.
        string outputPath = @"C:\Docs\Extracted.docx";

        // Load the source document. This uses the Document(string) constructor – a permitted lifecycle rule.
        Document sourceDoc = new Document(sourcePath);

        // Example extraction: take pages 1‑2 (zero‑based indices) from the source document.
        // The ExtractPages method returns a new Document containing only the specified pages.
        Document extractedDoc = sourceDoc.ExtractPages(0, 1);

        // Configure save options for DOCX output.
        // OoxmlSaveOptions is the appropriate SaveOptions class for DOCX files.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            // Preserve the current field results; do not recalculate them on save.
            UpdateFields = false,

            // Optional: omit the Aspose.Words generator name from the saved file.
            ExportGeneratorName = false
        };

        // Save the extracted document using the Save(string, SaveOptions) overload.
        // This follows the provided save rule and ensures the document is written as DOCX.
        extractedDoc.Save(outputPath, saveOptions);
    }
}
