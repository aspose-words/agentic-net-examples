using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportOutlinesToPdf
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PDF save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Show the outline (bookmarks / headings) pane when the PDF is opened.
        saveOptions.PageMode = PdfPageMode.UseOutlines;

        // Export bookmarks to the outline at level 1 (top level).
        saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;

        // Include headings up to level 3 in the outline.
        saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;

        // Optional: create missing outline levels so that gaps in heading levels are represented.
        saveOptions.OutlineOptions.CreateMissingOutlineLevels = true;

        // Save the document as PDF with the configured outline options.
        doc.Save("OutputWithOutlines.pdf", saveOptions);
    }
}
