// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

public static class MarkdownPrinter
{
    // Prints a Markdown file to the default printer.
    public static void PrintMarkdown(string markdownFilePath)
    {
        // Load the Markdown file. Aspose.Words automatically detects the format from the extension.
        Document doc = new Document(markdownFilePath);

        // Rebuild the page layout so that pagination and fields are up‑to‑date before printing.
        doc.UpdatePageLayout();

        // Print the whole document using the default printer.
        doc.Print();
    }

    // Prints a Markdown file to a specific printer and page range.
    public static void PrintMarkdown(string markdownFilePath, string printerName, int fromPage, int toPage)
    {
        Document doc = new Document(markdownFilePath);
        doc.UpdatePageLayout();

        // Configure printer settings.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = printerName,
            PrintRange = PrintRange.SomePages,
            FromPage = fromPage,
            ToPage = toPage
        };

        // Print using the specified settings and give the job a friendly name.
        doc.Print(printerSettings, "Markdown Document");
    }
}
