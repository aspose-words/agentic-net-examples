using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder where the split PDF files will be saved.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create a sample large Word document with a known number of pages.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Define how many pages the sample document will have.
        const int totalPages = 200; // Adjust as needed for testing.

        for (int i = 1; i <= totalPages; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert a page break after each page except the last one.
            if (i < totalPages)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure the layout is up‑to‑date so that PageCount is accurate.
        sourceDoc.UpdatePageLayout();
        int pageCount = sourceDoc.PageCount;

        // -----------------------------------------------------------------
        // Split the document into 50‑page chunks and save each chunk as PDF.
        // -----------------------------------------------------------------
        const int chunkSize = 50;
        int chunkIndex = 0;
        for (int start = 0; start < pageCount; start += chunkSize)
        {
            int pagesInChunk = Math.Min(chunkSize, pageCount - start);
            // ExtractPages uses zero‑based page indices.
            Document chunk = sourceDoc.ExtractPages(start, pagesInChunk);

            string pdfPath = Path.Combine(outputFolder, $"Chunk_{++chunkIndex}.pdf");
            chunk.Save(pdfPath, SaveFormat.Pdf);
        }

        // -----------------------------------------------------------------
        // Validation: ensure the expected number of PDF files were created.
        // -----------------------------------------------------------------
        int expectedFiles = (int)Math.Ceiling((double)pageCount / chunkSize);
        int actualFiles = Directory.GetFiles(outputFolder, "Chunk_*.pdf").Length;

        if (actualFiles != expectedFiles)
            throw new InvalidOperationException(
                $"Expected {expectedFiles} PDF files, but found {actualFiles}.");

        // The program finishes automatically; no user interaction required.
    }
}
