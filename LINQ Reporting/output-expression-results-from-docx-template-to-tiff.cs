using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class DocxToTiff
{
    static void Main()
    {
        // Paths to the input DOCX template and the output TIFF file.
        string dataDir = @"C:\Data\";
        string templatePath = Path.Combine(dataDir, "Template.docx");
        string outputPath = Path.Combine(dataDir, "Result.tiff");

        // Load the DOCX template.
        Document doc = new Document(templatePath);

        // Evaluate all fields (expressions) in the document so their results are rendered.
        doc.UpdateFields();

        // Configure image save options for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // NOTE: The MultiPageLayout property does not exist in Aspose.Words.
        // TIFF images are saved as multi‑page by default when using ImageSaveOptions.
        // If you need to control the page range, you can set PageIndex and PageCount.
        // options.PageIndex = 0;               // start from first page
        // options.PageCount = doc.PageCount;   // render all pages

        // Optional: set resolution (dpi) and image size if needed.
        options.Resolution = 300;                     // 300 dpi
        options.ImageSize = new Size(2480, 3508);      // A4 at 300 dpi

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);
    }
}
