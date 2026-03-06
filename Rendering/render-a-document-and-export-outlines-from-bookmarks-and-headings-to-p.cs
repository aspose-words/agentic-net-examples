using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document that contains headings and bookmarks.
        Document doc = new Document("MyDir/DocumentWithBookmarksAndHeadings.docx");

        // Create PDF save options to control outline generation.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Show the outline (navigation pane) when the PDF is opened.
        saveOptions.PageMode = PdfPageMode.UseOutlines;

        // Export all bookmarks at outline level 1.
        saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;

        // Include headings up to level 3 in the PDF outline.
        saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;

        // Expand the first two outline levels by default.
        saveOptions.OutlineOptions.ExpandedOutlineLevels = 2;

        // Save the document as PDF with the configured outline options.
        doc.Save("ArtifactsDir/DocumentWithOutlines.pdf", saveOptions);
    }
}
