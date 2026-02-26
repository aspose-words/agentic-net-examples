// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the Markdown file that contains the report.
        const string markdownFile = "Report.md";

        // Load the Markdown document. The LoadOptions tells Aspose.Words to treat the file as Markdown.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Markdown);
        Document doc = new Document(markdownFile, loadOptions);

        // If the document contains fields (e.g., DATE, PAGE), update them before printing.
        doc.UpdateFields();

        // Rebuild the page layout so that pagination and page‑related fields are correct.
        doc.UpdatePageLayout();

        // Print the document to the default printer.
        doc.Print();

        // Example: print to a specific printer with custom printer settings.
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Choose any installed printer; here we pick the first one.
            PrinterName = GetFirstInstalledPrinter(),
            PrintRange = PrintRange.AllPages
        };

        // Print using the specified settings and give the job a friendly name.
        doc.Print(printerSettings, "Markdown Report");
    }

    // Helper method that returns the name of the first installed printer.
    static string GetFirstInstalledPrinter()
    {
        foreach (string printer in PrinterSettings.InstalledPrinters)
            return printer;

        throw new InvalidOperationException("No printers are installed on this system.");
    }
}
