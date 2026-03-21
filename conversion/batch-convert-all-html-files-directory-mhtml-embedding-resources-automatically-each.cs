using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Use directories relative to the executable location.
        string baseDir = AppContext.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputHtml");
        string outputDir = Path.Combine(baseDir, "OutputMhtml");

        // Ensure both directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // If there are no HTML files, create a simple one for demonstration.
        if (Directory.GetFiles(inputDir, "*.html").Length == 0)
        {
            string samplePath = Path.Combine(inputDir, "sample.html");
            File.WriteAllText(samplePath, "<html><body><h1>Hello, World!</h1></body></html>");
        }

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

        Console.WriteLine($"Conversion complete. MHTML files are located in: {outputDir}");
    }
}
