using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source RTF file and the destination PDF file.
        string rtfPath = "input.rtf";
        string pdfPath = "output.pdf";

        // Load the RTF document using RtfLoadOptions.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document document = new Document(rtfPath, loadOptions);

        // Extract plain‑text content from the RTF file.
        PlainTextDocument plainText = new PlainTextDocument(rtfPath, loadOptions);
        string extractedText = plainText.Text;
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);

        // Save the document as PDF using PdfSaveOptions.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        document.Save(pdfPath, saveOptions);
    }
}
