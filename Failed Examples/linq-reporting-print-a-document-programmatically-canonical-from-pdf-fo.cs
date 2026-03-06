// ALL ATTEMPTS FAILED. Below is the last generated code.

using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

public class PdfPrintExample
{
    // Prints the entire PDF document to the default printer.
    public static void PrintPdf(string pdfFilePath)
    {
        // Load the PDF file into an Aspose.Words Document.
        Document doc = new Document(pdfFilePath);

        // Rebuild the page layout to ensure accurate pagination.
        doc.UpdatePageLayout();

        // Send the document to the default printer.
        doc.Print();
    }

    // Prints a range of pages from the PDF to a specific printer.
    public static void PrintPdfWithSettings(string pdfFilePath, string printerName, int fromPage, int toPage)
    {
        // Load the PDF file.
        Document doc = new Document(pdfFilePath);
        doc.UpdatePageLayout();

        // Configure printer settings.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = printerName,
            PrintRange = PrintRange.SomePages,
            FromPage = fromPage,
            ToPage = toPage
        };

        // Print using the specified settings.
        doc.Print(printerSettings);
    }

    // Prints the PDF using AsposeWordsPrintDocument for advanced tracking or preview.
    public static void PrintPdfWithPrintDocument(string pdfFilePath, string printerName)
    {
        // Load the PDF.
        Document doc = new Document(pdfFilePath);
        doc.UpdatePageLayout();

        // Create the Aspose.Words print document wrapper.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

        // Assign printer settings.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = printerName,
            PrintRange = PrintRange.AllPages
        };
        printDoc.PrinterSettings = printerSettings;

        // Optionally cache printer settings to speed up the first print call.
        printDoc.CachePrinterSettings();

        // Print the document.
        printDoc.Print();
    }
}
