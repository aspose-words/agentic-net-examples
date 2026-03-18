using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

class RetainOrientationWhenSplitting
{
    static void Main()
    {
        // Load the source document.
        Document srcDoc = new Document(@"Input\SourceDocument.docx");

        // Ensure the page layout information is up‑to‑date.
        srcDoc.UpdatePageLayout();

        // Folder where the split parts will be saved.
        string outFolder = @"Output\SplitParts";
        Directory.CreateDirectory(outFolder);

        // Iterate through each page of the source document.
        for (int pageIndex = 0; pageIndex < srcDoc.PageCount; pageIndex++)
        {
            // Extract a single page into a new document.
            // PageExtractOptions keeps page numbering and fields consistent.
            PageExtractOptions extractOptions = new PageExtractOptions();
            Document partDoc = srcDoc.ExtractPages(pageIndex, 1, extractOptions);

            // Determine the original orientation of the page we have just extracted.
            PageInfo pageInfo = srcDoc.GetPageInfo(pageIndex);
            Orientation originalOrientation = pageInfo.Landscape
                ? Orientation.Landscape
                : Orientation.Portrait;

            // Apply the original orientation to the first (and only) section of the part document.
            // This guarantees that landscape pages stay landscape after splitting.
            partDoc.Sections[0].PageSetup.Orientation = originalOrientation;

            // Save the part document. Here we use HTML format, but any format can be used.
            string partPath = Path.Combine(outFolder, $"Part_{pageIndex + 1}.html");
            partDoc.Save(partPath, SaveFormat.Html);
        }
    }
}
