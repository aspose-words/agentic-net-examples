// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class EnableLigaturesAndSavePdf
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\OutputDocument.pdf";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Enable standard ligatures for the whole document.
        // This setting affects how glyphs are rendered when the document is saved.
        doc.Font.Ligatures = FontLigatures.Standard;

        // Create PDF save options (default settings are sufficient for ligatures).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
