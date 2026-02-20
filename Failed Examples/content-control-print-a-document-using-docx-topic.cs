// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class PrintDocxExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Sample.docx");

        // Create the Aspose.Words implementation of PrintDocument.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

        // Optional: configure printer settings (e.g., page range, printer name).
        // PrintDialog printDlg = new PrintDialog();
        // if (printDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        // {
        //     printDoc.PrinterSettings = printDlg.PrinterSettings;
        // }

        // Print the document using the default printer.
        printDoc.Print();

        // Output the number of pages printed in color (useful for diagnostics).
        Console.WriteLine($"Color pages printed: {printDoc.ColorPagesPrinted}");
    }
}
