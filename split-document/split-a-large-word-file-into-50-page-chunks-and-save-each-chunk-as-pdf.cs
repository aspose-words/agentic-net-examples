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

        // Create a sample large document with a known number of pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        const int totalPages = 200; // Example size; adjust as needed.
        for (int i = 1; i <= totalPages; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < totalPages)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure the layout is up‑to‑date so PageCount is accurate.
        sourceDoc.UpdatePageLayout();
        int pageCount = sourceDoc.PageCount;

        const int chunkSize = 50; // Number of pages per PDF chunk.

        int chunkIndex = 1;
        for (int start = 0; start < pageCount; start += chunkSize)
        {
            int pagesRemaining = pageCount - start;
            int count = Math.Min(chunkSize, pagesRemaining);

            // Extract the required page range into a new document.
            Document chunk = sourceDoc.ExtractPages(start, count);

            // Save the chunk as PDF.
            string outPath = Path.Combine(outputDir, $"Chunk_{chunkIndex}.pdf");
            chunk.Save(outPath, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create PDF chunk: {outPath}");

            chunkIndex++;
        }

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine($"Document split into {chunkIndex - 1} PDF files in \"{outputDir}\".");
    }
}
