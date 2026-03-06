using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\ResultDocument.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // (Optional) If you need to customize PDF output, create PdfSaveOptions.
        // Here we use default options, but the object can be configured as required.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the save options.
        doc.Save(outputPath, pdfOptions);
    }
}
