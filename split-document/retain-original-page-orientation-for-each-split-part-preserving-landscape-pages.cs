using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Define folders for input and output.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains both portrait and
        //    landscape pages (two sections with different orientations).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section – portrait orientation (default).
        builder.Writeln("Page 1 – Portrait");
        builder.InsertBreak(BreakType.PageBreak);

        // Second section – landscape orientation.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Page 2 – Landscape");
        builder.InsertBreak(BreakType.PageBreak);

        // Third section – back to portrait.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("Page 3 – Portrait");

        // Save the source document.
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and split it page‑by‑page.
        //    Each extracted page retains its original page setup,
        //    including orientation.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        int pageCount = loadedDoc.PageCount;

        for (int i = 0; i < pageCount; i++)
        {
            // Determine the orientation of the current page before extraction.
            bool isLandscape = loadedDoc.GetPageInfo(i).Landscape;

            // Extract a single page (zero‑based start index, count = 1).
            Document pageDoc = loadedDoc.ExtractPages(i, 1);

            // Save the extracted page.
            string pageFileName = $"Page_{i + 1}.docx";
            string pagePath = Path.Combine(artifactsDir, pageFileName);
            pageDoc.Save(pagePath);

            // -----------------------------------------------------------------
            // 3. Validate that the saved file exists and that its orientation
            //    matches the original page's orientation.
            // -----------------------------------------------------------------
            if (!File.Exists(pagePath))
                throw new InvalidOperationException($"Failed to create split file: {pageFileName}");

            Document verifyDoc = new Document(pagePath);
            bool extractedIsLandscape = verifyDoc.Sections[0].PageSetup.Orientation == Orientation.Landscape;

            if (isLandscape != extractedIsLandscape)
                throw new InvalidOperationException($"Orientation mismatch in {pageFileName}");
        }

        // All pages have been split and validated successfully.
        Console.WriteLine($"Document split into {pageCount} parts. Files are located in: {artifactsDir}");
    }
}
