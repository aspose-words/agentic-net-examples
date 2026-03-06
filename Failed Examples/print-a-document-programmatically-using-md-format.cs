// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the Markdown file to be printed.
        // Aspose.Words automatically detects the .md format.
        string markdownPath = @"C:\Docs\Sample.md";

        // Load the Markdown document.
        Document doc = new Document(markdownPath);

        // Optional: configure printer settings (e.g., select a specific printer).
        // If you want to use the default printer, you can omit this block.
        PrinterSettings printerSettings = new PrinterSettings();
        // printerSettings.PrinterName = "Your Printer Name";

        // Print the document using the specified printer settings.
        // Uncomment the line below to use custom settings.
        // doc.Print(printerSettings);

        // Print using the default printer.
        doc.Print();

        // Inform the user that the print job has been sent.
        Console.WriteLine("Print job submitted successfully.");
    }
}
