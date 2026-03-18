using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class InsertPageBreakBeforeRegions
{
    static void Main()
    {
        // Load the source document that contains mail‑merge regions.
        Document doc = new Document("TemplateWithRegions.docx");

        // Get the hierarchy of all mail‑merge regions in the document.
        MailMergeRegionInfo rootInfo = doc.MailMerge.GetRegionsHierarchy();

        // Recursively insert a page break before each region start.
        InsertBreakBeforeRegion(doc, rootInfo.Regions);

        // Update fields if needed.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("TemplateWithRegions_PageBreaks.docx");
    }

    // Recursively processes a list of regions.
    private static void InsertBreakBeforeRegion(Document doc, IList<MailMergeRegionInfo> regions)
    {
        foreach (MailMergeRegionInfo region in regions)
        {
            // The start of the region is marked by a MERGEFIELD with the TableStart tag.
            // Move the builder to the start field's start node and insert a page break.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(region.StartField.Start);
            builder.InsertBreak(BreakType.PageBreak);

            // Process any nested regions.
            if (region.Regions != null && region.Regions.Count > 0)
                InsertBreakBeforeRegion(doc, region.Regions);
        }
    }
}
