using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Folder for output PDF chunks.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample large document (e.g., 200 pages).
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            const int totalSamplePages = 200;
            for (int i = 1; i <= totalSamplePages; i++)
            {
                builder.Writeln($"This is page {i}.");
                if (i < totalSamplePages)
                    builder.InsertBreak(BreakType.PageBreak);
            }

            // Ensure layout is up‑to‑date so PageCount is accurate.
            sourceDoc.UpdatePageLayout();
            int totalPages = sourceDoc.PageCount;

            // Split the document into 50‑page chunks.
            const int chunkSize = 50;
            int chunkCount = (totalPages + chunkSize - 1) / chunkSize;

            for (int i = 0; i < chunkCount; i++)
            {
                // ExtractPages uses zero‑based page index.
                int startPageZero = i * chunkSize; // zero‑based start page
                int pagesToExtract = Math.Min(chunkSize, totalPages - startPageZero);

                // Extract the required page range.
                Document chunk = sourceDoc.ExtractPages(startPageZero, pagesToExtract);

                // Save each chunk as a PDF file.
                string chunkPath = Path.Combine(outputDir, $"Chunk_{i + 1}.pdf");
                chunk.Save(chunkPath, SaveFormat.Pdf);
            }

            // Validate that all expected PDF files were created.
            for (int i = 0; i < chunkCount; i++)
            {
                string expectedPath = Path.Combine(outputDir, $"Chunk_{i + 1}.pdf");
                if (!File.Exists(expectedPath))
                    throw new InvalidOperationException($"Expected output file not found: {expectedPath}");
            }
        }
    }
}
