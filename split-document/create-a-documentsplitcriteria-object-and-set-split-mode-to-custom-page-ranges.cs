using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with enough content to span multiple pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;

        // Insert many lines; Aspose.Words will paginate automatically.
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln($"Line {i + 1}");
        }

        // Ensure layout information is up‑to‑date before extracting pages.
        doc.UpdatePageLayout();

        // Define the custom page ranges we want to split: 1‑2 and 3‑4.
        int[][] pageRanges = new int[][]
        {
            new int[] { 1, 2 },
            new int[] { 3, 4 }
        };

        // Prepare the output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save each page range as a separate HTML file.
        for (int i = 0; i < pageRanges.Length; i++)
        {
            int startPage = pageRanges[i][0];
            int endPage = pageRanges[i][1];
            int pageCount = endPage - startPage + 1;

            // Extract the required pages into a new document.
            Document splitDoc = doc.ExtractPages(startPage, pageCount);

            // Configure HTML save options (default options are sufficient here).
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Build the output file name.
            string fileName = i == 0 ? "SplitDocument.html" : $"SplitDocument_{i}.html";
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the split document.
            splitDoc.Save(outputPath, saveOptions);
        }

        // Validate that the primary output file was created.
        string primaryFile = Path.Combine(outputDir, "SplitDocument.html");
        if (!File.Exists(primaryFile))
            throw new Exception("The primary split HTML file was not created.");

        // Validate that at least one additional split file exists.
        string additionalFile = Path.Combine(outputDir, "SplitDocument_1.html");
        if (!File.Exists(additionalFile))
            throw new Exception("Expected additional split HTML file was not created.");

        Console.WriteLine("Document split completed successfully.");
    }
}
