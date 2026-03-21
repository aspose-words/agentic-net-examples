using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class RetainOrientationWhenSplitting
{
    static void Main()
    {
        // Ensure the input folder exists.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        Directory.CreateDirectory(inputFolder);

        // Path to the source document.
        string sourcePath = Path.Combine(inputFolder, "SourceDocument.docx");

        // If the source document does not exist, create a sample one with mixed orientations.
        if (!File.Exists(sourcePath))
        {
            Document sample = new Document();

            // First section – portrait.
            Section portraitSection = sample.Sections[0];
            portraitSection.PageSetup.Orientation = Orientation.Portrait;
            DocumentBuilder builder = new DocumentBuilder(sample);
            builder.Writeln("This is a portrait page.");
            builder.InsertBreak(BreakType.PageBreak);

            // Second section – landscape.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            Section landscapeSection = sample.Sections[1];
            landscapeSection.PageSetup.Orientation = Orientation.Landscape;
            builder.Writeln("This is a landscape page.");

            sample.Save(sourcePath);
        }

        // Load the source document.
        Document srcDoc = new Document(sourcePath);

        // Ensure the page layout information is up‑to‑date.
        srcDoc.UpdatePageLayout();

        // Folder where the split parts will be saved.
        string outFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output", "SplitParts");
        Directory.CreateDirectory(outFolder);

        // Iterate through each page of the source document.
        for (int pageIndex = 0; pageIndex < srcDoc.PageCount; pageIndex++)
        {
            // Extract a single page into a new document.
            PageExtractOptions extractOptions = new PageExtractOptions();
            Document partDoc = srcDoc.ExtractPages(pageIndex, 1, extractOptions);

            // Determine the original orientation of the page we have just extracted.
            var pageInfo = srcDoc.GetPageInfo(pageIndex);
            Orientation originalOrientation = pageInfo.Landscape ? Orientation.Landscape : Orientation.Portrait;

            // Apply the original orientation to the first (and only) section of the part document.
            partDoc.Sections[0].PageSetup.Orientation = originalOrientation;

            // Save the part document as HTML.
            string partPath = Path.Combine(outFolder, $"Part_{pageIndex + 1}.html");
            partDoc.Save(partPath, SaveFormat.Html);
        }

        Console.WriteLine("Splitting completed successfully.");
    }
}
