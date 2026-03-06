// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsPrintingDemo
{
    public class WordmlPrinter
    {
        /// <summary>
        /// Prints a WORDML (WordprocessingML) document directly from its XML string.
        /// The method loads the XML into an Aspose.Words Document, updates the layout,
        /// and sends it to the default printer.
        /// </summary>
        /// <param name="wordmlContent">The complete WORDML XML string.</param>
        public static void PrintFromWordml(string wordmlContent)
        {
            if (string.IsNullOrEmpty(wordmlContent))
                throw new ArgumentException("WORDML content cannot be null or empty.", nameof(wordmlContent));

            // Load the WORDML XML into a MemoryStream.
            // Document(Stream) constructor automatically detects the format (WORDML in this case).
            using (MemoryStream stream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(stream))
            {
                writer.Write(wordmlContent);
                writer.Flush();
                stream.Position = 0; // Reset position for reading.

                // Create the Document object from the stream.
                Document doc = new Document(stream);

                // Ensure the page layout is up‑to‑date before printing.
                doc.UpdatePageLayout();

                // Print the document to the default printer.
                doc.Print();
            }
        }

        /// <summary>
        /// Prints a WORDML document using custom printer settings (e.g., specific printer,
        /// page range, or document name). Demonstrates the use of AsposeWordsPrintDocument.
        /// </summary>
        /// <param name="wordmlContent">The WORDML XML string.</param>
        /// <param name="printerName">Name of the printer to use. If null, the default printer is used.</param>
        /// <param name="fromPage">First page to print (1‑based). Use 0 to print all pages.</param>
        /// <param name="toPage">Last page to print (1‑based). Use 0 to print all pages.</param>
        public static void PrintFromWordmlWithSettings(string wordmlContent, string printerName = null, int fromPage = 0, int toPage = 0)
        {
            if (string.IsNullOrEmpty(wordmlContent))
                throw new ArgumentException("WORDML content cannot be null or empty.", nameof(wordmlContent));

            using (MemoryStream stream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(stream))
            {
                writer.Write(wordmlContent);
                writer.Flush();
                stream.Position = 0;

                Document doc = new Document(stream);
                doc.UpdatePageLayout();

                // Create a PrintDocument implementation that knows how to render Aspose.Words documents.
                AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

                // Configure printer settings if a specific printer or page range is required.
                PrinterSettings settings = new PrinterSettings();

                if (!string.IsNullOrEmpty(printerName))
                    settings.PrinterName = printerName;

                if (fromPage > 0 && toPage >= fromPage)
                {
                    settings.PrintRange = PrintRange.SomePages;
                    settings.FromPage = fromPage;
                    settings.ToPage = toPage;
                }

                printDoc.PrinterSettings = settings;

                // Optional: cache printer settings to reduce first‑print latency.
                printDoc.CachePrinterSettings();

                // Print the document.
                printDoc.Print();
            }
        }

        // Example usage.
        public static void Main()
        {
            // Sample WORDML content (normally you would load this from a file or other source).
            string sampleWordml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:body>
                    <w:p><w:r><w:t>Hello from WORDML!</w:t></w:r></w:p>
                    <w:sectPr/>
                  </w:body>
                </w:document>";

            // Print using the default printer.
            PrintFromWordml(sampleWordml);

            // Print using a specific printer and page range (if the document had multiple pages).
            // Replace "Your Printer Name" with an actual installed printer name.
            PrintFromWordmlWithSettings(sampleWordml, printerName: "Your Printer Name", fromPage: 1, toPage: 1);
        }
    }
}
