using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input directory containing HTML files.
        string inputDir = @"C:\InputHtml";

        // Output directory where MHTML files will be saved.
        string outputDir = @"C:\OutputMhtml";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Get all *.html files in the input directory (non‑recursive).
        string[] htmlFiles = Directory.GetFiles(inputDir, "*.html");

        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML file into an Aspose.Words Document.
            Document doc = new Document(htmlPath);

            // Configure save options for MHTML.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs so that all resources (images, fonts, CSS) are embedded
                // in the MHTML package and referenced correctly by mail agents or browsers.
                ExportCidUrlsForMhtmlResources = true,

                // Export font resources as separate files inside the MHTML package.
                ExportFontResources = true,

                // Keep the default behavior of not embedding images as Base64;
                // they will be stored as separate MIME parts inside the MHTML.
                ExportImagesAsBase64 = false
            };

            // Build the output file name with .mht extension.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string mhtmlPath = Path.Combine(outputDir, fileNameWithoutExt + ".mht");

            // Save the document as MHTML using the configured options.
            doc.Save(mhtmlPath, saveOptions);
        }
    }
}
