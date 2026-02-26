// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsParagraphDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX file. The constructor automatically detects the format.
            Document doc = new Document("Template.docx");

            // Replace placeholder text in the whole document.
            // Example: replace {{CustomerName}} with actual name.
            doc.Range.Replace("{{CustomerName}}", "John Doe");

            // Save the modified document to a new file.
            doc.Save("Report.docx");

            // Print the document programmatically without showing any UI.
            // Create the Aspose.Words print document wrapper.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

            // Configure printer settings (optional). Here we use the default printer.
            PrinterSettings printerSettings = new PrinterSettings();
            printerSettings.PrintRange = PrintRange.AllPages;
            printDoc.PrinterSettings = printerSettings;

            // Print the document.
            printDoc.Print();

            // Optionally, you can also display a print dialog to let the user choose printer and page range.
            // Uncomment the following lines to use a dialog.
            /*
            PrintDialog printDialog = new PrintDialog
            {
                AllowSomePages = true,
                PrinterSettings = { MinimumPage = 1, MaximumPage = doc.PageCount, FromPage = 1, ToPage = doc.PageCount }
            };

            if (printDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                printDoc.PrinterSettings = printDialog.PrinterSettings;
                printDoc.Print();
            }
            */
        }
    }
}
