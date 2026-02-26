using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

namespace AsposeWordsPdfRendering
{
    // Custom warning collector that implements IWarningCallback.
    public class HandleDocumentWarnings : IWarningCallback
    {
        // Collection that stores all captured warnings.
        public WarningInfoCollection Warnings { get; } = new WarningInfoCollection();

        // This method is called by Aspose.Words whenever a warning occurs.
        public void Warning(WarningInfo info)
        {
            // Capture only formatting‑loss warnings (you can adjust the filter as needed).
            if (info.WarningType == WarningType.MinorFormattingLoss)
            {
                Console.WriteLine($"Warning: {info.Description}");
                Warnings.Warning(info);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths to the input DOCX file and the output PDF file.
            string inputPath = @"C:\Docs\Input.docx";
            string outputPath = @"C:\Docs\Output.pdf";

            // Load the source document.
            Document doc = new Document(inputPath);

            // Create a PdfSaveOptions instance to control PDF rendering.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Example: configure metafile rendering to use vector with fallback.
            MetafileRenderingOptions metafileOptions = new MetafileRenderingOptions
            {
                RenderingMode = MetafileRenderingMode.VectorWithFallback,
                EmulateRasterOperations = false
            };
            pdfOptions.MetafileRenderingOptions = metafileOptions;

            // Example: enable high‑quality rendering.
            pdfOptions.UseHighQualityRendering = true;

            // Attach the warning callback before any operation that may generate warnings.
            HandleDocumentWarnings warningCallback = new HandleDocumentWarnings();
            doc.WarningCallback = warningCallback;

            // Save the document as PDF using the configured options.
            doc.Save(outputPath, pdfOptions);

            // After saving, you can inspect the collected warnings.
            Console.WriteLine($"Total warnings captured: {warningCallback.Warnings.Count}");
        }
    }
}
