// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

namespace AsposeWordsPrintDotExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOT (Word template) file.
            // The file can be a .dot, .dotx or .dotm template.
            const string templatePath = @"C:\Templates\ReportTemplate.dot";

            // Load the DOT template into a Document object.
            // This uses the Document(string) constructor – the standard load rule.
            Document doc = new Document(templatePath);

            // OPTIONAL: If you need to specify a particular printer or printer settings,
            // create and configure a PrinterSettings instance.
            // Here we demonstrate printing to the default printer without any special settings.
            //PrinterSettings printerSettings = new PrinterSettings
            //{
            //    // Example: print only pages 1 to 2.
            //    PrintRange = PrintRange.SomePages,
            //    FromPage = 1,
            //    ToPage = 2
            //};

            // Print the document.
            // The Print() method prints to the default printer.
            // If you want to use custom settings, uncomment the line below and comment the simple Print() call.
            // doc.Print(printerSettings);
            doc.Print();

            // Inform the user that the print job has been sent.
            Console.WriteLine("Document has been sent to the printer.");
        }
    }
}
