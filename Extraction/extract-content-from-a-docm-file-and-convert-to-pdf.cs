using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string docmPath = @"C:\Docs\source.docm";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\result.pdf";

        // Load the DOCM document. The constructor automatically detects the format.
        Document document = new Document(docmPath);

        // (Optional) Extract plain‑text content from the document.
        // string plainText = document.GetText();

        // Create PDF save options – can be customized if needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the specified options.
        document.Save(pdfPath, pdfOptions);
    }
}
