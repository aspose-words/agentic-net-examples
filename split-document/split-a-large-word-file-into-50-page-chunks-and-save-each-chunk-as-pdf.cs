using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for output PDF chunks.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with many pages (e.g., 200 pages).
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
            // Aspose.Words uses zero‑based page indices for ExtractPages.
            int startPageZeroBased = i * chunkSize;                     // first page index of this chunk
            int pagesToExtract = Math.Min(chunkSize, totalPages - startPageZeroBased);

            // Extract the page range into a new document.
            Document chunk = sourceDoc.ExtractPages(startPageZeroBased, pagesToExtract);

            // Save the chunk as PDF.
            string pdfPath = Path.Combine(outputDir, $"Chunk_{i + 1}.pdf");
            chunk.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF chunk: {pdfPath}");
        }

        // Indicate completion.
        Console.WriteLine($"Document split into {chunkCount} PDF chunk(s) in folder: {outputDir}");
    }
}
