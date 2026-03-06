// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load a DOCX template that contains reporting tags.
        Document doc = new Document("Template.docx");

        // Example data source – can be any .NET object, DataSet, etc.
        var dataSource = new { Name = "John Doe", Address = "123 Main St" };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "ds");

        // Save the generated report (optional, demonstrates the save lifecycle rule).
        doc.Save("Report.docx");

        // -----------------------------------------------------------------
        // 1. Print directly to the default printer.
        // -----------------------------------------------------------------
        doc.Print();

        // -----------------------------------------------------------------
        // 2. Print to a specific printer by name.
        // -----------------------------------------------------------------
        if (PrinterSettings.InstalledPrinters.Count > 0)
        {
            string printerName = PrinterSettings.InstalledPrinters[0];
            doc.Print(printerName);
        }

        // -----------------------------------------------------------------
        // 3. Print using custom PrinterSettings (e.g., a page range).
        // -----------------------------------------------------------------
        PrinterSettings settings = new PrinterSettings
        {
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = Math.Min(2, doc.PageCount) // ensure we don't exceed page count
        };
        doc.Print(settings, "MyReport");

        // -----------------------------------------------------------------
        // 4. Print via a PrintDialog – user selects printer and options.
        // -----------------------------------------------------------------
        using (PrintDialog printDlg = new PrintDialog())
        {
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            if (printDlg.ShowDialog() == DialogResult.OK)
            {
                // AsposeWordsPrintDocument integrates the document with .NET printing.
                AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
                awPrintDoc.PrinterSettings = printDlg.PrinterSettings;
                awPrintDoc.ColorMode = ColorPrintMode.GrayscaleAuto;
                awPrintDoc.CachePrinterSettings(); // improves first‑print performance
                awPrintDoc.Print();
            }
        }

        // -----------------------------------------------------------------
        // 5. Show a Print Preview dialog and then print.
        // -----------------------------------------------------------------
        PrintPreviewDialog previewDlg = new PrintPreviewDialog();
        previewDlg.Show(); // bring the preview window to the front

        using (PrintDialog previewPrintDlg = new PrintDialog())
        {
            previewPrintDlg.AllowSomePages = true;
            previewPrintDlg.PrinterSettings.MinimumPage = 1;
            previewPrintDlg.PrinterSettings.MaximumPage = doc.PageCount;
            previewPrintDlg.PrinterSettings.FromPage = 1;
            previewPrintDlg.PrinterSettings.ToPage = doc.PageCount;

            if (previewPrintDlg.ShowDialog() != DialogResult.OK)
                return;

            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = previewPrintDlg.PrinterSettings;
            awPrintDoc.ColorMode = ColorPrintMode.GrayscaleAuto;
            awPrintDoc.CachePrinterSettings();

            // Hide the preview while we refresh it with the Aspose print document.
            previewDlg.Hide();
            previewDlg.PrintPreviewControl.InvalidatePreview();
            previewDlg.Document = awPrintDoc;
            previewDlg.ShowDialog();

            // Finally, send the document to the printer.
            awPrintDoc.Print();
        }
    }
}
