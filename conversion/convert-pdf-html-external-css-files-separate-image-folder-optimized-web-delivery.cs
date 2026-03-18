using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToHtmlConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path where the HTML file will be saved.
        string htmlPath = @"C:\Output\sample.html";

        // Determine folder for the external CSS file (same folder as HTML).
        string cssFolder = Path.GetDirectoryName(htmlPath);
        string cssFile = Path.Combine(cssFolder, "sample.css");

        // Folder for extracted images.
        string imagesFolder = Path.Combine(cssFolder, "images");

        // Ensure the images folder is clean and exists.
        if (Directory.Exists(imagesFolder))
            Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export CSS to an external stylesheet.
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = cssFile,

            // Export images as separate files into the images folder.
            ImagesFolder = imagesFolder,
            ExportImagesAsBase64 = false,

            // Optional: make the generated HTML more readable.
            PrettyFormat = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, saveOptions);
    }
}
