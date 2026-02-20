using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputDocPath = @"C:\Docs\sample.doc";

        // Path where the resulting PDF will be saved.
        string outputPdfPath = @"C:\Docs\sample.pdf";

        // Load the DOC file into a Document object. The constructor automatically detects the format.
        Document document = new Document(inputDocPath);

        // (Optional) Extract plain‑text content from the document.
        // PlainTextDocument plainText = new PlainTextDocument(inputDocPath);
        // Console.WriteLine(plainText.Text);

        // Save the loaded document as PDF using default PDF save options.
        document.Save(outputPdfPath, SaveFormat.Pdf);

        // If custom PDF options are required, use PdfSaveOptions instead:
        // PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // document.Save(outputPdfPath, pdfOptions);
    }
}
