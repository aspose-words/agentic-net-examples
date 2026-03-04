// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsMarkdownPrint
{
    class Program
    {
        static void Main()
        {
            // Path to the Markdown file that will be converted to a Word document.
            const string markdownPath = @"C:\Docs\Report.md";

            // Load the Markdown file into an Aspose.Words Document.
            // The LoadOptions constructor with LoadFormat.Markdown tells Aspose.Words to interpret the file as Markdown.
            var loadOptions = new LoadOptions(LoadFormat.Markdown);
            Document doc = new Document(markdownPath, loadOptions);

            // Ensure the layout is up‑to‑date before printing (optional but recommended).
            doc.UpdatePageLayout();

            // Print the whole document to the default printer.
            // If you need to specify a printer or printer settings, use the overloads that accept PrinterSettings.
            doc.Print();

            // Example of printing to a specific printer with a page range:
            //PrinterSettings settings = new PrinterSettings
            //{
            //    PrinterName = "Your Printer Name",
            //    PrintRange = PrintRange.SomePages,
            //    FromPage = 1,
            //    ToPage = 2
            //};
            //doc.Print(settings);
        }
    }
}
