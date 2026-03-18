using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class HtmlToMhtmlBatchConverter
{
    /// <summary>
    /// Converts all *.html files in <paramref name="inputFolder"/> to MHTML files in <paramref name="outputFolder"/>.
    /// All linked resources (images, fonts, CSS) are embedded automatically.
    /// </summary>
    public static void ConvertFolder(string inputFolder, string outputFolder)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Enumerate all HTML files in the source folder (non‑recursive).
        foreach (string htmlPath in Directory.EnumerateFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Prepare save options for MHTML.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs so that resources are referenced correctly inside the MHTML package.
                ExportCidUrlsForMhtmlResources = true,
                // Export font files as separate resources (they will be embedded in the MHTML container).
                ExportFontResources = true,
                // Export images as separate MIME parts (default behavior). No need for Base64 embedding.
                ExportImagesAsBase64 = false,
                // Keep the generated HTML tidy.
                PrettyFormat = true
            };

            // Build the output file name with .mht extension.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string mhtPath = Path.Combine(outputFolder, fileNameWithoutExt + ".mht");

            // Save the document as MHTML using the configured options.
            doc.Save(mhtPath, saveOptions);
        }
    }

    // Example entry point.
    static void Main()
    {
        // Folder containing source HTML files.
        string sourceFolder = @"C:\InputHtml";

        // Folder where the resulting MHTML files will be placed.
        string targetFolder = @"C:\OutputMhtml";

        ConvertFolder(sourceFolder, targetFolder);

        Console.WriteLine("Batch conversion completed.");
    }
}
