// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsPrintPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file.
            const string pdfPath = @"C:\Docs\Sample.pdf";

            // Load the PDF file into an Aspose.Words Document.
            // The Document constructor automatically detects the format.
            Document doc = new Document(pdfPath);

            // Ensure the layout is up‑to‑date before printing.
            // This is required if the document has been modified after loading.
            doc.UpdatePageLayout();

            // Option 1: Simple print using the default printer.
            // doc.Print();

            // Option 2: Print using a specific printer and page range.
            // Create a PrinterSettings object to control the print job.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example: print pages 1 through 3.
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 3,

                // Replace with the name of an installed printer if needed.
                // PrinterName = "Your Printer Name"
            };

            // Use AsposeWordsPrintDocument for richer control (e.g., color mode, progress tracking).
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
            printDoc.PrinterSettings = printerSettings;

            // Optional: print in grayscale to save ink.
            printDoc.ColorMode = ColorPrintMode.GrayscaleAuto;

            // Cache printer settings to reduce the first‑print latency.
            printDoc.CachePrinterSettings();

            // Send the document to the printer.
            printDoc.Print();

            Console.WriteLine("Print job submitted successfully.");
        }
    }
}
