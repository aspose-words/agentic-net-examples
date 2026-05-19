using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a sample document with six pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 6; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 6)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define the custom page ranges we want to export:
        // pages 1‑3 (zero‑based 0‑2) and pages 5‑6 (zero‑based 4‑5).
        var pageRanges = new (int start, int end)[]
        {
            (0, 2), // pages 1‑3
            (4, 5)  // pages 5‑6
        };

        // Extract each range and save it as a separate HTML file.
        for (int i = 0; i < pageRanges.Length; i++)
        {
            int start = pageRanges[i].start;
            int count = pageRanges[i].end - start + 1; // inclusive range

            // Extract the required pages into a new document.
            Document part = doc.ExtractPages(start, count);

            // Save the part. No splitting is required because we already have the exact range.
            string partFile = Path.Combine(outputDir, $"SplitDocument-{i + 1:D2}.html");
            part.Save(partFile, new HtmlSaveOptions());
        }

        // Verify that at least one split part was created.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument-*.html");
        if (splitFiles.Length == 0)
            throw new InvalidOperationException("No split document parts were created.");

        // List the created files.
        foreach (var file in splitFiles)
            Console.WriteLine($"Created: {Path.GetFileName(file)}");
    }
}
