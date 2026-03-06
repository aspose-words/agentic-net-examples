// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Rendering;

public class HtmlPrintExample
{
    /// <summary>
    /// Loads an HTML document from the specified path and prints it.
    /// </summary>
    /// <param name="htmlPath">Full path to the HTML file.</param>
    /// <param name="printerName">
    /// Optional printer name. If null or empty, the default printer is used.
    /// </param>
    public static void PrintHtmlDocument(string htmlPath, string printerName = null)
    {
        // Load the HTML file into an Aspose.Words Document.
        // The constructor automatically detects the format (HTML) from the file extension.
        Document doc = new Document(htmlPath);

        // If a specific printer is required, use the overload that accepts a printer name.
        // Otherwise, the document is sent to the default printer.
        if (!string.IsNullOrEmpty(printerName))
        {
            doc.Print(printerName);
        }
        else
        {
            doc.Print();
        }
    }

    /// <summary>
    /// Demonstrates printing an HTML document using a custom PrinterSettings object.
    /// This gives more control over page range, copies, etc.
    /// </summary>
    /// <param name="htmlPath">Full path to the HTML file.</param>
    /// <param name="printerName">Name of the printer to use.</param>
    public static void PrintHtmlWithSettings(string htmlPath, string printerName)
    {
        // Load the HTML document.
        Document doc = new Document(htmlPath);

        // Configure printer settings.
        PrinterSettings settings = new PrinterSettings
        {
            PrinterName = printerName,
            // Example: print only the first two pages.
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = 2,
            // Example: print two copies.
            Copies = 2
        };

        // Use AsposeWordsPrintDocument for richer printing events (optional).
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc)
        {
            PrinterSettings = settings
        };

        // Print the document.
        printDoc.Print();
    }

    // Example usage.
    public static void Main()
    {
        // Path to the HTML file to be printed.
        string htmlFile = @"C:\Docs\SampleReport.html";

        // Print to the default printer.
        PrintHtmlDocument(htmlFile);

        // Print to a specific printer with custom settings.
        string targetPrinter = "Microsoft Print to PDF";
        PrintHtmlWithSettings(htmlFile, targetPrinter);
    }
}
