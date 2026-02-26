// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Sample HTML that represents the report.
        string html = @"
            <html>
                <head><title>Sample Report</title></head>
                <body>
                    <h1>Sales Report</h1>
                    <p>Generated on: " + DateTime.Now.ToString("yyyy-MM-dd") + @"</p>
                    <table border='1' cellpadding='5'>
                        <tr><th>Product</th><th>Quantity</th><th>Price</th></tr>
                        <tr><td>Apple</td><td>10</td><td>$1.00</td></tr>
                        <tr><td>Banana</td><td>5</td><td>$0.50</td></tr>
                    </table>
                </body>
            </html>";

        // Convert the HTML string to a memory stream.
        using (MemoryStream htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html)))
        {
            // LoadOptions tells Aspose.Words that the incoming stream is HTML.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

            // Load the HTML into a Document object.
            Document doc = new Document(htmlStream, loadOptions);

            // Ensure fields (if any) are up‑to‑date and layout is calculated.
            doc.UpdateFields();
            doc.UpdatePageLayout();

            // ---------- Printing the document ----------
            // Print to the default printer.
            doc.Print();

            // Print to a specific printer with custom settings.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example printer name; replace with an installed printer on your machine.
                PrinterName = "Microsoft Print to PDF",
                PrintRange = PrintRange.AllPages
            };
            // The second overload allows you to specify a document name that appears in the print queue.
            doc.Print(printerSettings, "HTML_Report");

            // ---------- Optional: Save to PDF for verification ----------
            // Save the rendered document as PDF.
            doc.Save("HTML_Report.pdf", SaveFormat.Pdf);
        }
    }
}
