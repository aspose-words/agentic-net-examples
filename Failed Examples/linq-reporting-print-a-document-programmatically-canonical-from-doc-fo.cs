// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string docPath = @"C:\Docs\Sample.doc";

        // Print to the default printer.
        PrintDocument(docPath);

        // Uncomment the line below to print to a specific printer.
        // PrintDocument(docPath, "My Printer Name");
    }

    /// <summary>
    /// Loads a DOC file and sends it to a printer.
    /// </summary>
    /// <param name="docPath">Full path of the DOC file.</param>
    /// <param name="printerName">
    /// Optional printer name. If null or empty the default printer is used.
    /// </param>
    static void PrintDocument(string docPath, string printerName = null)
    {
        // Load the document from file (load rule).
        Document doc = new Document(docPath);

        // Rebuild the page layout so that pagination is current (layout rule).
        doc.UpdatePageLayout();

        if (string.IsNullOrEmpty(printerName))
        {
            // Print to the default printer (print rule).
            doc.Print();
        }
        else
        {
            // Use AsposeWordsPrintDocument for explicit printer selection.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);

            // Configure printer settings (standard .NET PrinterSettings).
            PrinterSettings settings = new PrinterSettings
            {
                PrinterName = printerName,
                PrintRange = PrintRange.AllPages
            };
            awPrintDoc.PrinterSettings = settings;

            // Cache printer settings to reduce first‑print latency (optional).
            awPrintDoc.CachePrinterSettings();

            // Print using the configured settings.
            awPrintDoc.Print();
        }
    }
}
