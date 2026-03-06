// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsReporting
{
    public class DocumentPrinter
    {
        /// <summary>
        /// Loads a DOC file, updates its layout and prints it.
        /// </summary>
        /// <param name="docPath">Full path to the source .doc document.</param>
        /// <param name="printerName">
        /// Optional printer name. If null or empty the default printer is used.
        /// </param>
        public void PrintDocument(string docPath, string printerName = null)
        {
            // Load the existing DOC document (lifecycle rule: use Document constructor).
            Document doc = new Document(docPath);

            // Ensure the page layout is up‑to‑date before printing.
            // This is required if the document was modified after the last render.
            doc.UpdatePageLayout();

            // If a specific printer is requested, configure PrinterSettings.
            if (!string.IsNullOrEmpty(printerName))
            {
                // Create a PrinterSettings instance and assign the desired printer.
                PrinterSettings printerSettings = new PrinterSettings
                {
                    PrinterName = printerName,
                    // Example: print all pages; modify as needed.
                    PrintRange = PrintRange.AllPages
                };

                // Use the Aspose.Words implementation of PrintDocument for richer control.
                AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc)
                {
                    PrinterSettings = printerSettings
                };

                // Optional: cache printer settings to reduce first‑print latency.
                awPrintDoc.CachePrinterSettings();

                // Print the document using the configured printer.
                awPrintDoc.Print();
            }
            else
            {
                // No specific printer – print to the default printer.
                // This uses the Document.Print() overload (lifecycle rule: use Print method).
                doc.Print();
            }
        }
    }
}
