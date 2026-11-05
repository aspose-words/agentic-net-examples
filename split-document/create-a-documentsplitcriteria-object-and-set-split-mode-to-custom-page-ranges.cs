using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with enough content to span several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Add enough text to force a new page.
            builder.Writeln(new string('x', 2000));
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure layout is calculated so PageCount is accurate.
        doc.UpdatePageLayout();

        // Define custom page ranges for splitting: pages 1‑2 and 4‑5.
        string[] customRanges = { "1-2", "4-5" };

        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputFolder);

        // Save each custom range as a separate HTML file.
        int rangeIndex = 1;
        foreach (string range in customRanges)
        {
            // Parse the range (e.g., "1-2").
            string[] parts = range.Split('-');
            if (parts.Length != 2 ||
                !int.TryParse(parts[0], out int startPage) ||
                !int.TryParse(parts[1], out int endPage) ||
                startPage < 1 ||
                endPage < startPage)
            {
                throw new InvalidOperationException($"Invalid page range: '{range}'.");
            }

            // Ensure the requested pages exist in the source document.
            if (endPage > doc.PageCount)
                throw new InvalidOperationException($"Requested page {endPage} exceeds document page count ({doc.PageCount}).");

            int pageCount = endPage - startPage + 1;

            // Extract the required pages (zero‑based start index).
            Document extracted = doc.ExtractPages(startPage - 1, pageCount);

            // Save the extracted pages as HTML.
            string fileName = Path.Combine(outputFolder, $"output_range{rangeIndex}.html");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection
            };
            extracted.Save(fileName, saveOptions);
            rangeIndex++;
        }

        // Validate that the expected split files were created.
        string[] generatedFiles = Directory.GetFiles(outputFolder, "output_range*.html");
        if (generatedFiles.Length != customRanges.Length)
            throw new InvalidOperationException($"Expected {customRanges.Length} HTML files, but found {generatedFiles.Length}.");

        Console.WriteLine("Split operation completed successfully. Generated files:");
        foreach (string file in generatedFiles)
        {
            Console.WriteLine(Path.GetFileName(file));
        }
    }
}
