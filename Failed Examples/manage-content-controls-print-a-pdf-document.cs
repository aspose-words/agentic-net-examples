// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Rendering;

class PrintPdfDocument
{
    static void Main()
    {
        // Path to the PDF file that will be printed.
        string pdfPath = @"C:\Docs\Sample.pdf";

        // Load the PDF into an Aspose.Words Document.
        // PdfLoadOptions can be used to customize loading; here we use the default options.
        Document doc = new Document(pdfPath, new PdfLoadOptions());

        // Create the Aspose.Words print document wrapper.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

        // Optional: set how non‑colored pages are printed (e.g., force grayscale for such pages).
        printDoc.ColorMode = ColorPrintMode.GrayscaleAuto;

        // Print the document using the default printer settings.
        // This call sends the document directly to the printer.
        printDoc.Print();

        // Output a simple confirmation.
        Console.WriteLine("PDF document has been sent to the printer.");
    }
}
