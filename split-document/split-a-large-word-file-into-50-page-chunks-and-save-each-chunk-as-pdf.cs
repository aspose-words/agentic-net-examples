using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for output PDF parts.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample large document (e.g., 200 pages).
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        const int totalPages = 200;
        for (int i = 1; i <= totalPages; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < totalPages)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure layout is up‑to‑date so PageCount is accurate.
        sourceDoc.UpdatePageLayout();
        int pageCount = sourceDoc.PageCount;

        // Split the document into 50‑page chunks.
        const int chunkSize = 50;
        int chunkCount = (pageCount + chunkSize - 1) / chunkSize;

        for (int chunkIndex = 0; chunkIndex < chunkCount; chunkIndex++)
        {
            // 1‑based page numbers for readability.
            int startPage = chunkIndex * chunkSize + 1;
            int endPage = Math.Min(startPage + chunkSize - 1, pageCount);

            // Convert to zero‑based index and count required by ExtractPages.
            int zeroBasedIndex = startPage - 1;
            int pagesToExtract = endPage - startPage + 1;

            // Extract the page range into a new document.
            Document part = sourceDoc.ExtractPages(zeroBasedIndex, pagesToExtract);

            // Save the part as PDF.
            string partFileName = Path.Combine(outputDir, $"Part_{chunkIndex + 1}.pdf");
            part.Save(partFileName, SaveFormat.Pdf);

            // Validate that the file was created.
            if (!File.Exists(partFileName))
                throw new InvalidOperationException($"Failed to create PDF part: {partFileName}");
        }

        // Indicate successful completion.
        Console.WriteLine($"Document split into {chunkCount} PDF file(s) in folder: {outputDir}");
    }
}
