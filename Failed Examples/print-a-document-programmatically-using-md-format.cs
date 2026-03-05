// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the Markdown file that will be loaded into a Word document.
        string markdownFile = @"C:\Docs\sample.md";

        // Load the Markdown file. The LoadOptions constructor specifies the format.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Markdown);
        Document doc = new Document(markdownFile, loadOptions);

        // Print the document to the default printer.
        doc.Print();

        // If you need to print to a specific printer, uncomment the following lines:
        // string printerName = "Your Printer Name";
        // doc.Print(printerName);
    }
}
