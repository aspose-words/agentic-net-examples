using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentWarningCollector : IWarningCallback
{
    public List<WarningInfo> CollectedWarnings { get; } = new List<WarningInfo>();

    public void Warning(WarningInfo info)
    {
        // Store and display each warning that occurs during saving.
        CollectedWarnings.Add(info);
        Console.WriteLine($"Warning ({info.WarningType}): {info.Description}");
    }
}

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Attach a warning callback to capture any warnings generated while saving.
        DocumentWarningCollector warningCallback = new DocumentWarningCollector();
        doc.WarningCallback = warningCallback;

        // Create a PdfSaveOptions instance via the factory method (lifecycle rule).
        SaveOptions genericOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        PdfSaveOptions pdfOptions = (PdfSaveOptions)genericOptions;

        // Set compliance to PDF/A‑4 + PDF/UA‑2, which enforces both standards.
        pdfOptions.Compliance = PdfCompliance.PdfA4Ua2;

        // Save the document as a PDF while the warning callback records any issues.
        doc.Save("Result.pdf", pdfOptions);

        // Optional: process collected warnings after saving.
        Console.WriteLine($"Total warnings: {warningCallback.CollectedWarnings.Count}");
    }
}
