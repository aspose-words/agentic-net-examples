using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Implements a warning callback to capture formatting loss warnings during rendering.
    public class HandleDocumentWarnings : IWarningCallback
    {
        public WarningInfoCollection Warnings { get; } = new WarningInfoCollection();

        public void Warning(WarningInfo info)
        {
            // Capture only minor formatting loss warnings (e.g., unsupported metafile operations).
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
            // Load the source document.
            Document doc = new Document("InputDocument.docx");

            // Attach the warning callback before any layout or rendering occurs.
            var warningCallback = new HandleDocumentWarnings();
            doc.WarningCallback = warningCallback;

            // Create a PdfSaveOptions object using the provided factory method.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

            // Example: enable memory optimization for large documents.
            pdfOptions.MemoryOptimization = true;

            // Save the document as PDF while applying the warning handling.
            doc.Save("RenderedOutput.pdf", pdfOptions);

            // Output the collected warnings, if any.
            if (warningCallback.Warnings.Count > 0)
            {
                Console.WriteLine("\nCollected warnings:");
                foreach (WarningInfo info in warningCallback.Warnings)
                {
                    Console.WriteLine($"- {info.Description}");
                }
            }
            else
            {
                Console.WriteLine("\nNo warnings were generated during rendering.");
            }
        }
    }
}
