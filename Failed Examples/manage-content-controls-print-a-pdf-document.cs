// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the PDF file into an Aspose.Words Document.
        // The constructor automatically detects the file format.
        Document pdfDocument = new Document("C:\\Path\\To\\YourDocument.pdf");

        // Print the whole document to the default printer.
        pdfDocument.Print();

        // If you need to print to a specific printer, uncomment the following lines:
        // string printerName = "Your Printer Name";
        // pdfDocument.Print(printerName);
    }
}
