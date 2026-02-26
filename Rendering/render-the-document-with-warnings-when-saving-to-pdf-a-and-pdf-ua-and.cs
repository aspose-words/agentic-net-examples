using System;
using Aspose.Words;
using Aspose.Words.Saving;

class HandleDocumentWarnings : IWarningCallback
{
    // Called by Aspose.Words when a warning occurs during saving.
    public void Warning(WarningInfo info)
    {
        // Output warning type and description to the console.
        Console.WriteLine($"{info.WarningType}: {info.Description}");
    }
}

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Register the warning callback to capture any warnings while saving.
        doc.WarningCallback = new HandleDocumentWarnings();

        // ---------- Save as PDF/A-1b ----------
        PdfSaveOptions pdfAOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/A-1b (visual appearance preservation).
            Compliance = PdfCompliance.PdfA1b
        };
        doc.Save("Output_PdfA1b.pdf", pdfAOptions);

        // ---------- Save as PDF/UA-1 ----------
        PdfSaveOptions pdfUaOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/UA-1 (accessibility compliance).
            Compliance = PdfCompliance.PdfUa1
        };
        doc.Save("Output_PdfUa1.pdf", pdfUaOptions);
    }
}
