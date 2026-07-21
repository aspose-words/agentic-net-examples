using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with six pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 6; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 6)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the original document.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        doc.Save(sourcePath);

        // Define custom page ranges (zero‑based): pages 1‑3 (indices 0‑2) and pages 5‑6 (indices 4‑5).
        var customRanges = new (int start, int count)[]
        {
            (0, 3), // pages 1‑3
            (4, 2)  // pages 5‑6
        };

        // Extract each range and save it as a separate HTML file.
        for (int i = 0; i < customRanges.Length; i++)
        {
            var (start, count) = customRanges[i];
            Document part = doc.ExtractPages(start, count);
            string partPath = Path.Combine(outputDir, $"SplitCustom_part{i + 1}.html");
            part.Save(partPath, SaveFormat.Html);
        }

        // Validate that the expected split files were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitCustom_part*.html");
        if (splitFiles.Length != customRanges.Length)
            throw new InvalidOperationException($"Expected {customRanges.Length} split HTML files, but found {splitFiles.Length}.");

        Console.WriteLine($"Generated {splitFiles.Length} split HTML file(s) in '{outputDir}'.");
    }
}
