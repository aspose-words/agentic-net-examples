using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToTiffConverter
{
    static void Main()
    {
        // Path to the DOCX template.
        string templatePath = @"C:\Docs\Template.docx";

        // Path where the multi‑page TIFF will be saved.
        string outputPath = @"C:\Docs\Result.tiff";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Ensure the page layout is up‑to‑date before rendering.
        doc.UpdatePageLayout();

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render each page as a separate frame in the TIFF file.
            PageLayout = MultiPageLayout.TiffFrames(),

            // Optional: set resolution (dpi) for the output images.
            Resolution = 300
        };

        // Save the document as a multi‑page TIFF.
        doc.Save(outputPath, saveOptions);
    }
}
