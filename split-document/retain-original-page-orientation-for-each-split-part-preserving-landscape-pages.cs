using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for the sample document and the split results.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string outputDir = Path.Combine(artifactsDir, "SplitPages");

        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains both portrait and
        //    landscape pages. The first section uses the default portrait
        //    orientation, the second section is set to landscape.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First (portrait) page.
        builder.Writeln("This is a portrait page. It uses the default orientation.");

        // Insert a new section that starts on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Change the orientation of the current section to landscape.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is a landscape page. Its orientation is set to Landscape.");

        // Save the source document for reference.
        string sourcePath = Path.Combine(artifactsDir, "SampleDocument.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document page by page, preserving the original
        //    orientation of each page. The ExtractPages method keeps the
        //    page setup (including orientation) intact.
        // -----------------------------------------------------------------
        int pageCount = sourceDoc.PageCount;

        for (int i = 0; i < pageCount; i++)
        {
            // Extract a single page (zero‑based index) into a new document.
            Document pageDoc = sourceDoc.ExtractPages(i, 1);

            // Build the output file name.
            string outFile = Path.Combine(outputDir, $"Split_Page_{i + 1}.docx");

            // Save the extracted page.
            pageDoc.Save(outFile);
        }

        // -----------------------------------------------------------------
        // 3. Validate that each split file was created successfully.
        // -----------------------------------------------------------------
        for (int i = 0; i < pageCount; i++)
        {
            string outFile = Path.Combine(outputDir, $"Split_Page_{i + 1}.docx");
            if (!File.Exists(outFile))
                throw new Exception($"Expected split file not found: {outFile}");

            // Optional: verify that the orientation matches the original page.
            Document splitDoc = new Document(outFile);
            Orientation orientation = splitDoc.FirstSection.PageSetup.Orientation;
            Console.WriteLine($"Page {i + 1} saved. Orientation: {orientation}");
        }

        // All done.
        Console.WriteLine("Document splitting completed successfully.");
    }
}
