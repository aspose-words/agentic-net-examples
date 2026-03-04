// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // HTML source to be printed.
        string html = @"<html>
<head><title>Sample</title></head>
<body>
    <h1>Hello Aspose.Words</h1>
    <p>This document was generated from an HTML string and printed programmatically.</p>
</body>
</html>";

        // Convert the HTML string to a UTF‑8 stream.
        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            // Load the HTML into an Aspose.Words Document.
            // LoadOptions with LoadFormat.Html ensures the content is interpreted as HTML.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Document doc = new Document(htmlStream, loadOptions);

            // Rebuild the page layout so that pagination is correct before printing.
            doc.UpdatePageLayout();

            // Print the whole document to the default printer.
            doc.Print();

            // ------------------------------------------------------------
            // Example: print to a specific printer with a custom page range.
            // ------------------------------------------------------------
            if (PrinterSettings.InstalledPrinters.Count > 0)
            {
                PrinterSettings printerSettings = new PrinterSettings
                {
                    // Choose the first installed printer (replace with a specific name if required).
                    PrinterName = PrinterSettings.InstalledPrinters[0],

                    // Print only the first three pages (adjust as needed).
                    PrintRange = PrintRange.SomePages,
                    FromPage = 1,
                    ToPage = Math.Min(3, doc.PageCount)
                };

                // Use the Document.Print overload that accepts PrinterSettings.
                doc.Print(printerSettings);
            }
        }
    }
}
