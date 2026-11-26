using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample source document that contains both portrait and landscape sections.
        const string sourcePath = "Source.docx";
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure the layout is up‑to‑date so that PageCount is accurate.
        sourceDoc.UpdatePageLayout();

        // Prepare the output folder.
        const string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // Total number of pages in the source document.
        int totalPages = sourceDoc.PageCount;

        // Split the document page by page.
        // Aspose.Words uses zero‑based page indexes for ExtractPages.
        for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            // Extract a single page while preserving its formatting and orientation.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Save the extracted page as a separate DOCX file.
            string outPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.docx");
            pageDoc.Save(outPath, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create split file: {outPath}");
        }
    }

    // Generates a sample document containing portrait and landscape sections.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – portrait orientation.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("Section 1 – Portrait");
        builder.Writeln("Content on page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content on page 2.");

        // Second section – landscape orientation.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Section 2 – Landscape");
        builder.Writeln("Content on page 3.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content on page 4.");

        // Save the source document.
        doc.Save(path, SaveFormat.Docx);
    }
}
