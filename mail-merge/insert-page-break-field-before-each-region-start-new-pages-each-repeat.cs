using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class InsertPageBreakBeforeRegions
{
    static void Main()
    {
        // Create a simple template document with a mail‑merge region.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some content before the region.
        builder.Writeln("Document Header");

        // Insert the start of the mail‑merge region.
        builder.InsertField("MERGEFIELD TableStart:MyRegion");

        // Insert a table that will be repeated for each data row.
        builder.StartTable();
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Name");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Age");
        builder.EndRow();
        builder.EndTable();

        // Insert the end of the mail‑merge region.
        builder.InsertField("MERGEFIELD TableEnd:MyRegion");

        // Insert some content after the region.
        builder.Writeln("Document Footer");

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
