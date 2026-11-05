using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample large document (e.g., 200 pages).
        const int totalPagesToCreate = 200;
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= totalPagesToCreate; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert a page break after each page except the last one.
            if (i < totalPagesToCreate)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure layout is up‑to‑date so PageCount is accurate.
        sourceDoc.UpdatePageLayout();
        int totalPageCount = sourceDoc.PageCount;

        // Prepare output folder.
        string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // Split the document into 50‑page chunks and save each chunk as PDF.
        const int chunkSize = 50;
        int chunkIndex = 1;
        for (int startPageZeroBased = 0; startPageZeroBased < totalPageCount; startPageZeroBased += chunkSize)
        {
            int pagesToExtract = Math.Min(chunkSize, totalPageCount - startPageZeroBased);
            Document chunk = sourceDoc.ExtractPages(startPageZeroBased, pagesToExtract);
            string pdfPath = Path.Combine(outputFolder, $"Chunk_{chunkIndex}.pdf");
            chunk.Save(pdfPath); // Format inferred from .pdf extension.
            chunkIndex++;
        }

        // Validate that the expected number of PDF files were created.
        int expectedChunkCount = (int)Math.Ceiling((double)totalPageCount / chunkSize);
        string[] pdfFiles = Directory.GetFiles(outputFolder, "*.pdf");
        if (pdfFiles.Length != expectedChunkCount)
        {
            throw new InvalidOperationException(
                $"Expected {expectedChunkCount} PDF files, but found {pdfFiles.Length}.");
        }
    }
}
