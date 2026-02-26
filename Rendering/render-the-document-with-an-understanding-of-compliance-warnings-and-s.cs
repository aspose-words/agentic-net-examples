using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual path).
        Document doc = new Document("Input.docx");

        // Capture any compliance‑related warnings that Aspose.Words may emit.
        doc.WarningCallback = new WarningInfoCallback();

        // Configure PDF save options with the desired compliance level.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: PDF/A‑1b compliance – preserves visual appearance.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }

    // Simple implementation of IWarningCallback to output warnings to the console.
    private class WarningInfoCallback : IWarningCallback
    {
        public void Warning(WarningInfo info)
        {
            Console.WriteLine($"Warning: {info.WarningType} - {info.Description}");
        }
    }
}
