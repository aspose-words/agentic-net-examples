using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocmToPdfConverter
{
    static void Main()
    {
        // Path to the source DOCM file (macro‑enabled Word document)
        string sourcePath = @"C:\Input\SampleDocument.docm";

        // Path where the resulting PDF will be saved
        string outputPath = @"C:\Output\SampleDocument.pdf";

        // Load the DOCM file. The Document constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Optional: configure PDF save options (e.g., compliance level, image quality, etc.)
        // Here we use the default options; you can customize if needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF. The Save method determines the format from the file extension.
        doc.Save(outputPath, pdfOptions);
    }
}
