using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define directories for input and output.
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with multiple pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document (optional, for reference).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Create a DocumentSplitCriteria instance and set a split mode.
        // Here we use PageBreak as an example; custom page‑range splitting will be performed manually.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak;

        // Define custom page ranges (zero‑based indices). Example: pages 1‑2 and 4‑5.
        var customRanges = new (int start, int count)[]
        {
            (0, 2), // pages 1 and 2
            (3, 2)  // pages 4 and 5
        };

        // Extract each range and save as a separate document.
        for (int i = 0; i < customRanges.Length; i++)
        {
            var (start, count) = customRanges[i];
            Document part = sourceDoc.ExtractPages(start, count);
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            part.Save(partPath);
        }

        // Verify that the expected split files were created.
        for (int i = 1; i <= customRanges.Length; i++)
        {
            string partPath = Path.Combine(outputDir, $"Part_{i}.docx");
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split document not found: {partPath}");
        }

        // Optional: demonstrate saving the source document to HTML using the split criteria.
        // This will split the HTML output at each page break.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria
        };
        string htmlPath = Path.Combine(outputDir, "Source.html");
        sourceDoc.Save(htmlPath, htmlOptions);
    }
}
