using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Create a sample source document with enough content to span many pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Generate 200 pages of simple text.
        for (int i = 1; i <= 200; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert a page break after each page except the last.
            if (i < 200)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Determine total page count of the source document.
        int totalPages = sourceDoc.PageCount;

        // Split the document into 50‑page chunks and save each chunk as PDF.
        int chunkIndex = 1;
        for (int startPage = 1; startPage <= totalPages; startPage += 50)
        {
            // Number of pages to extract for this chunk.
            int pagesToExtract = Math.Min(50, totalPages - startPage + 1);

            // ExtractPages uses zero‑based page index, so subtract 1 from startPage.
            Document chunk = sourceDoc.ExtractPages(startPage - 1, pagesToExtract);

            string pdfPath = Path.Combine(outputFolder, $"Chunk_{chunkIndex}.pdf");
            chunk.Save(pdfPath, SaveFormat.Pdf);
            chunkIndex++;
        }

        // Validate that the expected PDF files were created.
        int expectedChunkCount = (totalPages + 49) / 50; // Ceiling division.
        string[] pdfFiles = Directory.GetFiles(outputFolder, "Chunk_*.pdf");
        if (pdfFiles.Length != expectedChunkCount)
            throw new InvalidOperationException($"Expected {expectedChunkCount} PDF files, but found {pdfFiles.Length}.");

        foreach (string file in pdfFiles)
        {
            if (!File.Exists(file))
                throw new FileNotFoundException($"Expected output file not found: {file}");
        }

        // Program completes without waiting for user input.
    }
}
