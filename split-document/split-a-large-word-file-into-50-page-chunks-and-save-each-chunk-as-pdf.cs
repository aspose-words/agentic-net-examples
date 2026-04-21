using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentIntoPdfChunks
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample large document with many pages.
        string sourcePath = Path.Combine(outputDir, "LargeDocument.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Generate 200 pages of sample content.
        const int totalSamplePages = 200;
        for (int i = 1; i <= totalSamplePages; i++)
        {
            builder.Writeln($"This is page {i} of the sample document.");
            if (i < totalSamplePages)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document (optional, useful for inspection).
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Ensure the document layout is up‑to‑date.
        sourceDoc.UpdatePageLayout();

        // Determine total pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Split size (pages per chunk).
        const int pagesPerChunk = 50;

        // Calculate how many chunks are needed.
        int chunkCount = (totalPages + pagesPerChunk - 1) / pagesPerChunk;

        for (int i = 0; i < chunkCount; i++)
        {
            // Zero‑based start page index.
            int startPageIndex = i * pagesPerChunk;
            // Number of pages for this chunk (may be less than 50 for the last chunk).
            int pageCount = Math.Min(pagesPerChunk, totalPages - startPageIndex);

            // Extract the required page range into a new document.
            Document chunk = sourceDoc.ExtractPages(startPageIndex, pageCount);

            // Save the chunk as PDF.
            string chunkPath = Path.Combine(outputDir, $"Chunk_{i + 1}.pdf");
            chunk.Save(chunkPath, SaveFormat.Pdf);

            // Validate that the file was created.
            if (!File.Exists(chunkPath))
                throw new InvalidOperationException($"Failed to create PDF chunk: {chunkPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine($"Document split into {chunkCount} PDF chunk(s) in folder: {outputDir}");
    }
}
