using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportOutlineToPdf
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PDF save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Show the outline (bookmarks / headings) pane when the PDF is opened.
        saveOptions.PageMode = PdfPageMode.UseOutlines;

        // Export all bookmarks at the first outline level.
        saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;

        // Include headings up to level 3 in the PDF outline.
        saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;

        // (Optional) Create missing outline levels so that gaps in heading levels are represented.
        saveOptions.OutlineOptions.CreateMissingOutlineLevels = true;

        // Save the document as PDF with the configured outline options.
        doc.Save("OutputDocument.pdf", saveOptions);
    }
}
