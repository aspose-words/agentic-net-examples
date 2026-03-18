using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source PDF document.
        Document doc = new Document("input.pdf");

        // Create image save options for JPEG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

        // Render all pages side‑by‑side in a horizontal layout.
        // The argument is the spacing (in points) between pages.
        options.PageLayout = MultiPageLayout.Horizontal(10);

        // Set a high JPEG quality (0‑100). 90 gives excellent quality with reasonable size.
        options.JpegQuality = 90;

        // Optional: improve rendering quality.
        options.UseHighQualityRendering = true;
        options.UseAntiAliasing = true;

        // Save the whole document as a single JPEG image.
        doc.Save("output.jpg", options);
    }
}
