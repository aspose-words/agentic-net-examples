using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the source document and output folder.
        string sourcePath = "source.docx";
        string outputFolder = "output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Create a sample source document with many pages (e.g., 200 pages).
        CreateSampleDocument(sourcePath, 200);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Total pages in the source document.
        int totalPages = sourceDoc.PageCount;
        int chunkSize = 50; // Number of pages per split chunk.

        // Split the document into 50‑page chunks and save each chunk as PDF.
        int partNumber = 1;
        for (int startPage = 1; startPage <= totalPages; startPage += chunkSize)
        {
            // Calculate how many pages to extract for this chunk.
            int pagesToExtract = Math.Min(chunkSize, totalPages - startPage + 1);

            // Aspose.Words uses zero‑based page indexes for ExtractPages.
            int zeroBasedStart = startPage - 1;

            // Extract the required pages.
            Document partDoc = sourceDoc.ExtractPages(zeroBasedStart, pagesToExtract);

            // Save the extracted part as PDF.
            string outputPath = Path.Combine(outputFolder, $"Part_{partNumber}.pdf");
            partDoc.Save(outputPath, SaveFormat.Pdf);

            partNumber++;
        }

        // Validate that the expected number of PDF files were created.
        int expectedFiles = (totalPages + chunkSize - 1) / chunkSize;
        int actualFiles = Directory.GetFiles(outputFolder, "*.pdf").Length;
        if (actualFiles != expectedFiles)
        {
            throw new InvalidOperationException(
                $"Expected {expectedFiles} PDF files, but found {actualFiles}.");
        }
    }

    // Helper method to create a sample document with a specified number of pages.
    private static void CreateSampleDocument(string filePath, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < pageCount)
            {
                // Insert a page break to start a new page.
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        doc.Save(filePath);
    }
}
