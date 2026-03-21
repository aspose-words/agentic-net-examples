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

    static void Main()
    {
        // Use temporary directories so the example works out‑of‑the‑box.
        string tempRoot = Path.Combine(Path.GetTempPath(), "HtmlToMhtmlDemo");
        string sourceFolder = Path.Combine(tempRoot, "InputHtml");
        string targetFolder = Path.Combine(tempRoot, "OutputMhtml");

        // Ensure clean state.
        if (Directory.Exists(tempRoot))
            Directory.Delete(tempRoot, recursive: true);

        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(targetFolder);

        // Create a simple HTML file for demonstration.
        string sampleHtmlPath = Path.Combine(sourceFolder, "sample.html");
        File.WriteAllText(sampleHtmlPath,
            @"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Sample</title>
    <style>body { font-family: Arial; }</style>
</head>
<body>
    <h1>Hello, World!</h1>
    <p>This is a sample HTML file.</p>
</body>
</html>");

        // Perform the batch conversion.
        ConvertFolder(sourceFolder, targetFolder);

        Console.WriteLine($"Batch conversion completed. Output files are located in: {targetFolder}");
    }
}
