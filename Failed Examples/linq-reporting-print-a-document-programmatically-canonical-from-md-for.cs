// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the Markdown file that contains the report.
        const string markdownFile = "Report.md";

        // Load the Markdown file into an Aspose.Words Document.
        // LoadOptions tells Aspose.Words to interpret the source as Markdown.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Markdown);
        Document doc = new Document(markdownFile, loadOptions);

        // Re‑calculate the page layout so that pagination is correct before printing.
        doc.UpdatePageLayout();

        // Print the entire document to the default printer.
        doc.Print();

        // ------------------------------------------------------------
        // Optional: print a specific page range using custom printer settings.
        // ------------------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrintRange = PrintRange.SomePages, // Enable page range selection.
            FromPage = 1,                      // First page to print (1‑based index).
            ToPage = 2                         // Last page to print.
        };

        // Print using the specified settings and give the job a recognizable name.
        doc.Print(printerSettings, "MarkdownReport_PrintJob");
    }
}
